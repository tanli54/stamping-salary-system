# 冲压车间工资录入系统

一个基于 Python + tkinter 的计件工资录入工具，支持产品管理、员工管理、工资录入、汇总查询。

## 功能
- 员工增删管理
- 产品名称+单价管理（支持关键词搜索）
- 工资录入：输入日期、产品、数量，自动计算工钱
- 工资汇总表导出
- 工资查询
- 适老化大字体界面

## 环境要求
- Windows 7 / 10 / 11
- Python 3.7+（开发）
- 依赖：pandas、openpyxl、pyinstaller

## 使用方法
1. 安装依赖：`pip install pandas openpyxl pyinstaller`
2. 准备数据文件（参考 `示例数据.xlsx` 格式）
3. 运行：`python 工资系统.py`
4. 打包：`pyinstaller --onefile --windowed --name="冲压工资系统" 工资系统.py`

## 数据说明
- 首次运行读取 `示例数据.xlsx`（或你的计数表）生成产品库和员工名单
- 之后数据保存在 `产品库.txt`、`员工名单.txt`、`工资记录.xlsx`
- 这些文件已加入 `.gitignore`，不会上传

## 注意
- 本仓库所有数据均为虚构示例，不含真实员工和产品信息
