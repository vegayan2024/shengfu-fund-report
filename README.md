# 晟孚资本 · 基金产品月报

## 净值数据更新规则

### 上传格式

将净值 Excel 文件放入 `data/` 目录，按以下命名规范：

```
data/
├── zhouqi_YYYYMM.xlsx    # 泽鑫周期
├── jiazhi_YYYYMM.xlsx    # 泽鑫价值
```

例如：`data/zhouqi_202608.xlsx`、`data/jiazhi_202608.xlsx`

### Excel 格式要求

| 日期 | 单位净值 |
|------|---------|
| 2024-07-05 | 0.204 |
| 2024-07-08 | 0.195 |
| ... | ... |

- **第一列**：日期（支持 `YYYY-MM-DD` 或 Excel 日期格式）
- **第二列**：单位净值（数字）
- 无需特定表头名称，脚本按列位置读取
- 支持 `.xlsx` 和 `.xls` 格式

### 更新步骤

```bash
# 1. 将 Excel 文件放入 data/ 目录
cp ~/Desktop/zhouqi_202608.xlsx data/

# 2. 提交并推送
git add data/
git commit -m "更新2026年8月净值数据"
git push origin main

# 3. GitHub Actions 自动处理
#    - 读取 Excel 文件
#    - 获取沪深300基准数据
#    - 更新 combined-data.json
#    - 自动部署到 GitHub Pages
```

### 本地测试

```bash
pip install xlrd akshare
python scripts/update_nav.py
```

### 数据处理流程

1. 读取 `data/` 目录下最新月份的 Excel 文件
2. 从 akshare 获取沪深300指数数据
3. 日期对齐：只保留基金和基准都有数据的日期
4. 归一化处理：起始日净值 = 1.0
5. 输出到 `zhouqi/combined-data.json` 和 `jiazhi/combined-data.json`

### 页面访问

- 首页：`https://vegayan2024.github.io/shengfu-fund-report/`
- 泽鑫周期：`https://vegayan2024.github.io/shengfu-fund-report/zhouqi/`
- 泽鑫价值：`https://vegayan2024.github.io/shengfu-fund-report/jiazhi/`
