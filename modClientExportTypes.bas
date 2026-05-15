Attribute VB_Name = "modClientExportTypes"
Option Explicit

'=========================================================
' Option types used across the add-in
'=========================================================
Public Enum eDataAction
    daClear = 0
    daFlatten = 1
End Enum

Public Enum eVisibilityAction
    vaPreserve = 0
    vaExpand = 1
End Enum

Public Enum eObjectPolicy
    opKeepAll = 0
    opKeepInsidePrintArea = 1
    opDeleteAll = 2
End Enum

Public Type ExportOptions
    GroupedDataAction As eDataAction
    GroupedVisibilityAction As eVisibilityAction

    HiddenDataAction As eDataAction
    HiddenVisibilityAction As eVisibilityAction

    imagePolicy As eObjectPolicy
    chartPolicy As eObjectPolicy
End Type
