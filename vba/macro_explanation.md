# VBA Macro Explanation — Slicer Connection Automation

This document explains the VBA macro used in the MIS Sales Performance Dashboard to automate slicer connections across multiple PivotTables.  
The macro improves user experience by allowing checkbox‑based control of slicer connections.

---

## Purpose of the Macro

The macro dynamically **adds or removes slicer connections** between the Region slicer and multiple PivotTables based on checkbox selections in the dashboard.

This allows the user to:

- Connect the slicer to selected PivotTables
- Disconnect the slicer from others
- Control dashboard views with a single click
- Avoid manual slicer connection steps

---

## How the Macro Works

The macro checks the value of specific cells (B1, E1, H1, K1) that act as **checkbox triggers**.

Each cell corresponds to a PivotTable:

| Cell | PivotTable  | Purpose                   |
| ---- | ----------- | ------------------------- |
| B1   | PivotTable1 | Connect/Disconnect slicer |
| E1   | PivotTable3 | Connect/Disconnect slicer |
| H1   | PivotTable4 | Connect/Disconnect slicer |
| K1   | PivotTable5 | Connect/Disconnect slicer |

If the cell value is **TRUE**, the slicer is connected.  
If the cell value is **FALSE**, the slicer is disconnected.

---

## Macro Code Used

```vba
Sub SlicerConnection()

If Sheet1.Range("B1").Value = True Then
    ActiveWorkbook.SlicerCaches("Slicer_Region").PivotTables.AddPivotTable ( _
        ActiveSheet.PivotTables("PivotTable1"))
Else
    ActiveWorkbook.SlicerCaches("Slicer_Region").PivotTables.RemovePivotTable ( _
        ActiveSheet.PivotTables("PivotTable1"))
End If

If Sheet1.Range("E1").Value = True Then
    ActiveWorkbook.SlicerCaches("Slicer_Region").PivotTables.AddPivotTable ( _
        ActiveSheet.PivotTables("PivotTable3"))
Else
    ActiveWorkbook.SlicerCaches("Slicer_Region").PivotTables.RemovePivotTable ( _
        ActiveSheet.PivotTables("PivotTable3"))
End If

If Sheet1.Range("H1").Value = True Then
    ActiveWorkbook.SlicerCaches("Slicer_Region").PivotTables.AddPivotTable ( _
        ActiveSheet.PivotTables("PivotTable4"))
Else
    ActiveWorkbook.SlicerCaches("Slicer_Region").PivotTables.RemovePivotTable ( _
        ActiveSheet.PivotTables("PivotTable4"))
End If

If Sheet1.Range("K1").Value = True Then
    ActiveWorkbook.SlicerCaches("Slicer_Region").PivotTables.AddPivotTable ( _
        ActiveSheet.PivotTables("PivotTable5"))
Else
    ActiveWorkbook.SlicerCaches("Slicer_Region").PivotTables.RemovePivotTable ( _
        ActiveSheet.PivotTables("PivotTable5"))
End If

End Sub
```
