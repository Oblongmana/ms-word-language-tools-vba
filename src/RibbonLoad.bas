Attribute VB_Name = "RibbonLoad"
'Adding refreshability to our ribbon
'cf. https://spreadsheetgurucourses.com/checkbox-control/
'NB: unused at present, decent likelihood it will be needed later

Option Private Module
Option Explicit

Public customRibbon As IRibbonUI

Sub languageToolsRibbonOnLoad(ribbon As IRibbonUI)
    Set customRibbon = ribbon
End Sub

Sub RefreshRibbon()
    customRibbon.Invalidate
End Sub
