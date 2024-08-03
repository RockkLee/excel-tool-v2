# Spec for excel tool
## Filter
- Supplier Name
- Raw Material
- Color

## Output
- First outputs:
  - Style
  - Component
- Final outputs:
  - Use the outputs of Style & Color to filter the final outputs
  - The final outputs should include all columns

## The Filename
- Filename: `{Supplier_names}-{Raw_materials}-{Colors}{num (if the current filename exists)}.xlsx`

## Export the project to exe file
```python
mkdir exe
cd exe
pyinstaller ../excel_tool.py ../lib.py ../dto.py --onefile -w
```

---

# Spec 2.0
- Add search bars in UI
- Replace the Component filter with the Color filter
