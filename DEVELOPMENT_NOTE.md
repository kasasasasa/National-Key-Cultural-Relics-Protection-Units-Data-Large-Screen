# Development Note: 文物主题大屏精简与视觉改造

## Goal

删除项目中企业数据大屏与招聘数据大屏相关内容，仅保留全国重点文物保护单位大屏，并让页面样式更贴近文物主题。

## Requirement / Assumptions

- 用户要求删除 `4600 万企业数据大屏可视化` 与 `厦门 10 万招聘数据大屏可视化` 相关内容。
- 删除范围包含项目内可见入口、路由与接口说明，避免继续暴露这两套演示数据入口。
- 保留首页 `/` 的文物数据大屏功能不变。
- 样式改造以现有布局和图表结构为基础，不做新的业务模块扩展。

## Business Boundary

- Context: Flask 单页大屏展示
- Business capability: 文物保护单位数据可视化展示
- Owns: 首页模板、主样式、首页接口与说明文档
- Does not own: Excel 数据源结构、ECharts 基础库、第三方地图资源

## Design Summary

- 后端删除 `/corp`、`/job` 与对应 `/api/corp`、`/api/job`。
- 前端刷新逻辑统一回落到首页 `/api/data`，移除对 `corp/job` 路径分支的依赖。
- `README.md` 移除两条演示入口和相关说明。
- 样式上改为深褐、金石、绢纸质感方向，强化展陈氛围，保留现有三栏布局和图表区域。

## Implementation Plan

1. 精简 `app.py` 路由与无用导入。
2. 更新 `templates/index.html` 中的数据刷新路径逻辑与页面文案细节。
3. 更新 `static/css/comon0.css`，调整背景、头图、面板、地图区与文字配色。
4. 更新 `README.md`，移除企业/招聘大屏说明。

## Verification Plan

- 启动 Flask 服务并访问 `http://127.0.0.1:5000/`。
- 确认首页正常展示、图表正常刷新、无空白页或脚本报错。
- 确认 `README.md` 不再出现 `corp` 与 `job` 两个演示入口。
- 手动检查整体视觉更偏文物展陈主题，无明显文字溢出和布局重叠。

## Risks / Non-goals

- 不处理编码历史遗留问题，除非本次改动直接涉及。
- 不重做图表配置结构，仅在现有视觉基础上增强主题气质。
- 不承诺删除所有历史演示数据文件，仅移除项目对外暴露与说明中的相关内容。

## Handoff Notes

- 若后续需要彻底移除 `static_data/corp.json`、`static_data/job.json` 和对应数据类，可在确认无复用需求后再做一次清理。
