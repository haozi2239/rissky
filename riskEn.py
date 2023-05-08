import csv
import xlsxwriter
import argparse

# 提取漏洞信息的函数
def extract_vulnerabilities(csv_file):
    critical_vulns = []
    high_vulns = []
    medium_vulns = []

    # 指定正确的编码格式打开文件
    with open(csv_file, 'r', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        for row in reader:
            if row['Risk'] == 'Critical':
                vuln = {
                    'IP': row['Host'],
                    'Vulnerability Name': row['Name'],
                    'Solution': row['Solution']
                }
                if row['Protocol'] == 'tcp':
                    critical_vulns.append(vuln)
                    
            elif row['Risk'] == 'High':
                vuln = {
                    'IP': row['Host'],
                    'Vulnerability Name': row['Name'],
                    'Solution': row['Solution']
                }
                if row['Protocol'] == 'tcp':
                    high_vulns.append(vuln)
                    
            elif row['Risk'] == 'Medium':
                vuln = {
                    'IP': row['Host'],
                    'Vulnerability Name': row['Name'],
                    'Solution': row['Solution']
                }
                if row['Protocol'] == 'tcp':
                    medium_vulns.append(vuln)

    return {'严重': critical_vulns, '高危': high_vulns, '中危': medium_vulns}

# 创建新的 ArgumentParser 对象，用于解析命令行参数
parser = argparse.ArgumentParser(description='提取 Nessus 报告中的漏洞信息并将其写入 Excel 工作表。')
parser.add_argument('csv_file', metavar='CSV_FILE', type=str, help='Nessus 报告的 CSV 文件路径')
parser.add_argument('-o', '--output', metavar='OUTPUT_FILE', type=str, default='nessus_vulnerabilities.xlsx', help='输出文件名')

# 解析命令行参数
args = parser.parse_args()

# 提取漏洞信息
vulnerabilities = extract_vulnerabilities(args.csv_file)

# 计算漏洞数量
total = 0
num_critical = len(vulnerabilities['严重'])
total += num_critical
num_high = len(vulnerabilities['高危'])
total += num_high
num_medium = len(vulnerabilities['中危'])
total += num_medium

# 创建 Excel 工作簿
workbook = xlsxwriter.Workbook(args.output)

# 添加工作表并设置标题
for risk, vulns in vulnerabilities.items():
    worksheet = workbook.add_worksheet(risk)
    worksheet.write('A1', 'IP')
    worksheet.write('B1', '漏洞名称')
    worksheet.write('C1', '修复建议')

    # 将漏洞信息写入工作表
    row = 1
    for vuln in vulns:
        worksheet.write(row, 0, vuln['IP'])
        worksheet.write(row, 1, vuln['Vulnerability Name'])
        worksheet.write(row, 2, vuln['Solution'])
        row += 1

# 添加报告快速统计信息到工作簿
worksheet = workbook.add_worksheet('报告统计')
bold = workbook.add_format({'bold': True})
worksheet.write('A1', '风险等级', bold)
worksheet.write('B1', '漏洞数量', bold)
worksheet.write('A2', '严重')
worksheet.write('B2', num_critical)
worksheet.write('A3', '高危')
worksheet.write('B3', num_high)
worksheet.write('A4', '中危')
worksheet.write('B4', num_medium)
worksheet.write('A5', '总计')
worksheet.write('B5', total)

# 设置工作表格式
worksheet.set_column('A:A', 20)
worksheet.set_column('B:B', 50)
worksheet.set_column('C:C', 80)

# 关闭 Excel 工作簿
workbook.close()

# 输出报告统计信息到控制台
print(f'总漏洞数量: {total}')
print(f'严重漏洞数量: {num_critical}')
print(f'高危漏洞数量: {num_high}')
print(f'中危漏洞数量: {num_medium}')
print(f'已经完成，请您直接打开输出的文件')
