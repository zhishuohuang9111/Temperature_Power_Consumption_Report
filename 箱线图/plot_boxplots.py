"""读取测试 Excel，按 Excel 的 Slot 编号生成箱线图。
运行：python3 plot_boxplots.py 05-34-39
依赖：openpyxl numpy matplotlib
"""
import argparse
import csv
from pathlib import Path
import numpy as np
import openpyxl
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('folder', nargs='?', default='05-34-39')
    parser.add_argument('--power-unit', default='raw value', help='确认单位后可设为 W')
    args = parser.parse_args()
    folder = Path(args.folder)
    files = list(folder.glob('*.xlsx'))
    if len(files) != 1:
        raise ValueError('文件夹必须恰好包含一份 xlsx')
    out = folder / '箱线图输出'
    out.mkdir(exist_ok=True)
    wb = openpyxl.load_workbook(files[0], read_only=True, data_only=True)
    collected = {}
    summary = []
    for name in ['Temp', 'DimmPower', 'AmtTemp']:
        sheet = wb[name]
        rows = iter(sheet.values)
        headers = next(rows)
        columns = {i: [] for i, h in enumerate(headers) if isinstance(h, str) and h.startswith('Slot')}
        first = last = None
        count = 0
        for row in rows:
            count += 1
            first = row[0] if first is None else first
            last = row[0]
            for i in columns:
                value = row[i]
                if isinstance(value, (int, float)) and np.isfinite(value):
                    columns[i].append(value)
        data = {headers[i]: np.asarray(values) for i, values in columns.items() if values}
        collected[name] = data
        print(name, 'rows=', count, 'time=', first, 'to', last)
        for slot, values in data.items():
            q1, median, q3 = np.percentile(values, [25, 50, 75])
            iqr = q3 - q1
            summary.append([name, slot, len(values), values.min(), q1, median, q3, values.max(),
                            int(((values < q1 - 1.5*iqr) | (values > q3 + 1.5*iqr)).sum())])
        print('active slots:', list(data), 'range:', min(v.min() for v in data.values()), max(v.max() for v in data.values()))
    wb.close()
    with (out / 'slot_statistics.csv').open('w', encoding='utf-8-sig', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(['sheet', 'excel_slot', 'n', 'min', 'q1', 'median', 'q3', 'max', 'outliers_1.5IQR'])
        writer.writerows(summary)
    plt.rcParams.update({'font.size': 10, 'axes.spines.top': False, 'axes.spines.right': False})
    fig, axes = plt.subplots(2, 1, figsize=(14, 9), constrained_layout=True)
    for ax, name, label, color in zip(axes, ['Temp', 'DimmPower'], ['Temperature (°C)', f'Power ({args.power_unit})'], ['#e88962', '#6fa8dc']):
        data = collected[name]
        slots = sorted(data, key=lambda s: int(s[4:]))
        bp = ax.boxplot([data[s] for s in slots], tick_labels=slots, patch_artist=True, whis=1.5,
                        flierprops={'marker': '.', 'markersize': 1.5, 'alpha': .25},
                        medianprops={'color': '#172b4d', 'linewidth': 1.6})
        for box in bp['boxes']:
            box.set_facecolor(color)
            box.set_alpha(.75)
        ax.set_ylabel(label)
        ax.grid(axis='y', alpha=.2)
        ax.set_title(f'{name}: distribution over all valid timestamps', loc='left')
        ax.set_xticks(range(1, len(slots)+1), [s + '\nn=' + str(len(data[s])) for s in slots], fontsize=8)
    fig.suptitle(f'{folder.name} — Temperature and power by Excel slot', fontsize=16)
    fig.savefig(out / 'temperature_power_by_slot.png', dpi=180)
    fig.savefig(out / 'temperature_power_by_slot.svg')
    plt.close(fig)
    print('Saved:', out)

if __name__ == '__main__':
    main()
