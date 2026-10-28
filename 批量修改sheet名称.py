# -*- coding: utf-8 -*-
"""
批量修改 Excel Sheet 名称工具

规则：
  - 映射表放在名为 Sheet1 的工作表中（可在下方 MAPPING_SHEET 修改）
  - A列 = 原sheet名称，B列 = 新sheet名称，从第2行开始（可在 START_ROW 修改）
  - 工具按 A 列名称找到对应的 sheet，改名为 B 列名称

用法：
  双击运行：会弹出文件选择窗口，选中 Excel 文件即可
  或拖拽：把 Excel 文件拖到 exe/脚本图标上
  或命令行：python 批量修改sheet名称.py "文件.xlsx"

保存前会自动在同目录生成一个带时间戳的备份文件。
"""
import os
import sys
import shutil
import datetime


def pause(msg='按回车退出...'):
    """等待用户回车；非交互环境（双击/管道）下直接跳过不报错"""
    try:
        input(msg)
    except (EOFError, KeyboardInterrupt):
        pass


def pick_file_dialog():
    """弹出文件选择窗口（不依赖拖拽），返回所选路径；取消/失败返回 None"""
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)  # 置顶，避免弹窗落到其他窗口后面
        path = filedialog.askopenfilename(
            title='请选择要批量修改 sheet 名称的 Excel 文件',
            filetypes=[('Excel 文件', '*.xlsx *.xlsm'), ('所有文件', '*.*')],
        )
        root.destroy()
        return path or None
    except Exception:
        return None


try:
    import openpyxl
except ImportError:
    print('缺少 openpyxl 库，请先执行: pip install openpyxl')
    pause()
    sys.exit(1)

# ================== 可配置项 ==================
MAPPING_SHEET = 'Sheet1'   # 存放映射表（A列原名/B列新名）的工作表名称
START_ROW = 2              # 映射数据起始行
# =============================================

INVALID_CHARS = set(':\\/?*[]')


def cell_to_name(v):
    """把单元格值规范成 sheet 名称字符串"""
    if v is None:
        return None
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = str(v).strip()
    return s if s else None


def validate_name(name):
    """校验新名称是否符合 Excel sheet 命名规则，返回错误原因或 None"""
    if not name:
        return '名称为空'
    if len(name) > 31:
        return f'名称超过31个字符（当前 {len(name)} 个）'
    bad = sorted({ch for ch in name if ch in INVALID_CHARS})
    if bad:
        return f'含非法字符 {bad}（不能含 : \\ / ? * [ ]）'
    if name.startswith("'") or name.endswith("'"):
        return "名称不能以单引号 ' 开头或结尾"
    return None


def read_mapping(ws):
    """从映射表读取 (原名称, 新名称) 列表"""
    mapping = []
    for row in ws.iter_rows(min_row=START_ROW, max_col=2, values_only=True):
        old, new = cell_to_name(row[0]), cell_to_name(row[1])
        if old is None and new is None:
            continue  # 跳过整行为空
        mapping.append((old, new))
    return mapping


def plan_renames(wb, mapping):
    """
    校验并生成重命名执行计划。
    返回 (todo, errors)：
      todo   = [(原名称, 新名称), ...] 可安全直接执行的
      errors = [(原名称, 新名称, 原因), ...] 有问题跳过的
    """
    errors = []
    todo = []

    existing = {name.casefold(): name for name in wb.sheetnames}
    will_rename = {}   # casefold(原名) -> 原名（这些名字会被释放掉）
    for old, new in mapping:
        if old is not None:
            will_rename[old.casefold()] = old

    seen_targets = {}  # casefold(新名) -> 原名（用于检测新名重复）
    for old, new in mapping:
        row_no = None
        if old is None:
            errors.append(('', new or '', 'A列（原名称）为空'))
            continue
        if new is None:
            errors.append((old, '', 'B列（新名称）为空'))
            continue
        if old == new:
            continue  # 名称没变，跳过即可，不算错误

        # 原名称必须存在（可能已被本表前面的行改过名，做一次动态检查）
        if old.casefold() not in existing:
            errors.append((old, new, '找不到该名称的sheet（可能已改过名或名称不符）'))
            continue

        # 新名称规则校验
        reason = validate_name(new)
        if reason:
            errors.append((old, new, f'新名称不合法：{reason}'))
            continue

        # 新名称重复（Excel sheet 名不区分大小写）
        if new.casefold() in seen_targets:
            errors.append((old, new, f'新名称与第 {seen_targets[new.casefold()]} 个映射重复'))
            continue

        # 新名称与其他不改名的现有 sheet 冲突
        if new.casefold() in existing and new.casefold() not in will_rename:
            errors.append((old, new, '新名称与现有其他sheet重名'))
            continue

        seen_targets[new.casefold()] = old
        todo.append((old, new))
        # 原名即将被释放，登记新名防止后续映射再撞上中间状态
        existing[new.casefold()] = new

    return todo, errors


def execute_renames(wb, todo):
    """
    执行重命名。对于“链式/环式”改名（如 A→B 且 B→C），
    先统一改成临时名再改最终名，避免中途重名冲突。
    返回成功列表 [(原名称, 新名称), ...]
    """
    todo_cf = {old.casefold(): new for old, new in todo}
    # 判断是否会撞中间状态：目标名恰好也是待改名的原名
    needs_temp = [old for old, new in todo if new.casefold() in todo_cf and new.casefold() != old.casefold()]

    # 第一阶段：有冲突风险的先改临时名
    used = {n.casefold() for n in wb.sheetnames}
    temp_map = {}  # 原名 -> 临时名
    for i, old in enumerate(needs_temp):
        tmp = f'~tmp{i}~'
        while tmp.casefold() in used:
            tmp += 'x'
        wb[old].title = tmp
        temp_map[old] = tmp

    # 第二阶段：改成最终名
    done = []
    for old, new in todo:
        src = temp_map.get(old, old)
        wb[src].title = new
        done.append((old, new))
    return done


def main():
    if len(sys.argv) > 1:
        path = sys.argv[1]
    else:
        # 双击运行：先弹文件选择窗口，取消或弹不出来再退回手动输入
        print('双击运行模式下，请在弹出的窗口中选择 Excel 文件...')
        path = pick_file_dialog()
        if not path:
            try:
                path = input('未选择文件。请输入 Excel 文件路径（可把文件拖进本窗口）：').strip().strip('"')
            except EOFError:
                print('未选择任何文件，程序退出。')
                pause()
                sys.exit(0)

    path = os.path.abspath(path)
    if not os.path.isfile(path):
        print(f'文件不存在：{path}')
        pause()
        sys.exit(1)

    print(f'正在读取：{path}')
    wb = openpyxl.load_workbook(path)

    if MAPPING_SHEET not in wb.sheetnames:
        print(f'错误：找不到映射工作表 “{MAPPING_SHEET}”（可修改脚本顶部的 MAPPING_SHEET）')
        pause()
        sys.exit(1)

    # A列可能是公式（如 WPS 的 _wpsfn.SHEETSNAME），
    # 默认模式只能读到公式对象，这里用 data_only 再读一份拿公式的计算结果（缓存值）
    wb_values = openpyxl.load_workbook(path, data_only=True, read_only=True)
    mapping = read_mapping(wb_values[MAPPING_SHEET])
    wb_values.close()
    if not mapping:
        print(f'错误：{MAPPING_SHEET} 中从第 {START_ROW} 行起没有读到任何映射数据')
        pause()
        sys.exit(1)

    print(f'共读到 {len(mapping)} 行映射，开始校验...\n')

    todo, errors = plan_renames(wb, mapping)

    for old, new, reason in errors:
        print(f'  [跳过] {old!r} -> {new!r}：{reason}')
    if errors:
        print(f'\n共 {len(errors)} 行有问题被跳过')

    if not todo:
        print('没有可执行的改名操作。')
        pause()
        sys.exit(0)

    # 备份
    stamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    root, ext = os.path.splitext(path)
    backup = f'{root}_{stamp}_备份{ext or ".xlsx"}'
    shutil.copy2(path, backup)
    print(f'已生成备份：{os.path.basename(backup)}')

    done = execute_renames(wb, todo)

    # 微信/QQ 接收的文件常带只读属性，保存前先去掉
    if not os.access(path, os.W_OK):
        try:
            os.chmod(path, 0o666)
            print('提示：原文件为只读，已自动去除只读属性')
        except OSError:
            pass

    try:
        wb.save(path)
    except PermissionError:
        print(f'\n保存失败：文件被占用，请先关闭 Excel/WPS 中的该文件再重试。'
              f'\n（改名操作尚未写入，原文件未受影响）')
        pause()
        sys.exit(1)

    print(f'\n完成：成功修改 {len(done)} 个 sheet 名称')
    for old, new in done:
        print(f'  [成功] {old}  ->  {new}')
    print(f'\n文件已保存：{path}')
    pause()


if __name__ == '__main__':
    main()
