# 发票整理 Skill

本项目以非破坏方式整理中国发票和报销票据。原始文件始终保留在 `input/`，内部计划位于 `.runtime/`，最终结果位于 `output/`。

## 五步使用

1. 把需要整理的原始票据放入 `input/`。
2. 只读扫描：

   ```powershell
   python scripts/scan_invoices.py
   ```

3. 查看 `.runtime/发票整理预览.md`，核对命名、顺序、字段和待确认项。
4. 明确确认后，使用预览中的令牌执行：

   ```powershell
   python scripts/organize_invoices.py --plan .runtime/.invoice-organization-plan.json --confirm <TOKEN>
   ```

5. 校验并查看结果：

   ```powershell
   python scripts/validate_organization.py --plan .runtime/.invoice-organization-plan.json
   ```

`output/` 只包含整理后的票据文件和 `发票整理清单.xlsx`。ZIP 只读处理并永久保留在 input；ZIP 中非 PDF 成员不会输出。

## 自定义目录

```powershell
python scripts/scan_invoices.py --input "D:\报销资料\input" --output "D:\报销资料\output"
```

`.runtime` 始终位于 Skill 项目根目录。input 与 output 不能相同或互相嵌套。

## 依赖

```powershell
python -m pip install -r requirements.txt
```
