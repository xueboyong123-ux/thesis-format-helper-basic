# Release Checklist

发布前请逐项检查。

- [ ] `git status` 是否干净
- [ ] 是否运行 `python wfp.py --test`
- [ ] 是否运行 `python wfp_cli.py test`
- [ ] 是否运行 `python -m py_compile wfp_gui.py wfp_config.py wfp_core.py wfp_tests.py wfp.py wfp_cli.py`
- [ ] README 是否完整
- [ ] LICENSE 是否保留
- [ ] 是否说明二次开发来源
- [ ] 是否没有个人路径
- [ ] 是否没有真实论文文件
- [ ] 是否没有 `__pycache__`
- [ ] 是否没有 `.pyc`
- [ ] 是否没有测试输出 docx
- [ ] 是否没有格式报告输出 txt
- [ ] 是否人工打开 GUI 测试
- [ ] 是否用测试 docx 跑过完整流程
- [ ] 是否检查 report 正常生成
