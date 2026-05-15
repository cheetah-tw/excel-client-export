VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmClientExportOptions 
   Caption         =   "UserForm1"
   ClientHeight    =   5625
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   11076
   OleObjectBlob   =   "frmClientExportOptions.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmClientExportOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Cancelled As Boolean

Private Sub OptionButton1_Click()

End Sub

Private Sub CommandButton1_Click()

End Sub

Private Sub Label1_Click()

End Sub

Private Sub UserForm_Initialize()
    Me.Caption = "Client Export Settings"
    Cancelled = True

    'Default starting selections for the dialog
    optGroupedFlatten.Value = True
    optGroupedKeep.Value = True

    optHiddenFlatten.Value = True
    optHiddenKeep.Value = True

    optImagesKeepInside.Value = True
    optChartsKeepInside.Value = True

    'Tooltips - Grouped
    optGroupedClear.ControlTipText = "Delete the contents of all grouped row/column cells while preserving formatting. Group structure stays unless you also choose Expand groups."
    optGroupedFlatten.ControlTipText = "Replace formulas in grouped row/column cells with their displayed values while preserving formatting."
    optGroupedKeep.ControlTipText = "Preserve the current grouped/collapsed display state."
    optGroupedExpand.ControlTipText = "Expand grouped rows/columns after processing so grouped sections are visible."

    'Tooltips - Hidden
    optHiddenClear.ControlTipText = "Delete the contents of manually hidden row/column cells while preserving formatting."
    optHiddenFlatten.ControlTipText = "Replace formulas in manually hidden row/column cells with values while preserving formatting."
    optHiddenKeep.ControlTipText = "Keep manually hidden rows/columns hidden after processing."
    optHiddenExpand.ControlTipText = "Unhide manually hidden rows/columns after processing."

    'Tooltips - Images
    optImagesKeepAll.ControlTipText = "Keep all imported pictures everywhere on the sheet."
    optImagesKeepInside.ControlTipText = "Delete pictures outside the sheet's print area and keep only those inside."
    optImagesDeleteAll.ControlTipText = "Delete all imported pictures regardless of location."

    'Tooltips - Charts
    optChartsKeepAll.ControlTipText = "Keep all chart objects everywhere on the sheet."
    optChartsKeepInside.ControlTipText = "Delete chart objects outside the sheet's print area and keep only those inside."
    optChartsDeleteAll.ControlTipText = "Delete all chart objects regardless of location."
End Sub

Private Sub cmdRun_Click()
    Cancelled = False
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    Cancelled = True
    Me.Hide
End Sub

Public Function GetOptions() As ExportOptions
    Dim o As ExportOptions

    'Grouped
    If optGroupedClear.Value Then
        o.GroupedDataAction = daClear
    Else
        o.GroupedDataAction = daFlatten
    End If

    If optGroupedExpand.Value Then
        o.GroupedVisibilityAction = vaExpand
    Else
        o.GroupedVisibilityAction = vaPreserve
    End If

    'Hidden
    If optHiddenClear.Value Then
        o.HiddenDataAction = daClear
    Else
        o.HiddenDataAction = daFlatten
    End If

    If optHiddenExpand.Value Then
        o.HiddenVisibilityAction = vaExpand
    Else
        o.HiddenVisibilityAction = vaPreserve
    End If

    'Images
    If optImagesKeepAll.Value Then
        o.imagePolicy = opKeepAll
    ElseIf optImagesKeepInside.Value Then
        o.imagePolicy = opKeepInsidePrintArea
    Else
        o.imagePolicy = opDeleteAll
    End If

    'Charts
    If optChartsKeepAll.Value Then
        o.chartPolicy = opKeepAll
    ElseIf optChartsKeepInside.Value Then
        o.chartPolicy = opKeepInsidePrintArea
    Else
        o.chartPolicy = opDeleteAll
    End If

    GetOptions = o
End Function
