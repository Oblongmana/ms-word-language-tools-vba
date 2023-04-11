Attribute VB_Name = "AutoCorrectInterfaceHandlers"
'Handling and menu creation for Quick Settings for AutoCorrect

Option Private Module
Option Explicit

'Callback for ltAutoCorrectMenu getContent
Sub ltAutoCorrectMenu_getContent(control As IRibbonControl, ByRef returnedVal)
    'Display and toggle various AutoCorrect settings.
    '  NB (copied from customUI XML comment)
    '  Note: this is dynamic so that we can signal to the user the state of the settings in a timely way (i.e. inside the menu!).
    '  Document and Application settings do not have any change events, so there's nothing we can hook onto to invalidate the ribbon.
    '  Other ribbon elements such as <menu> don't have any way of recalculating without full invalidation, so e.g. if you use a normal <menu>
    '  the values get cached the first time you click it, so if you then change a setting e.g. by editing the Styles["Normal"] template
    '  directly, then click the menu again, the value will be inaccurate! Whereas the dynamicMenu, using invalidateContentOnDrop, can
    '  recalculate every time it's clicked, and so display current values, regardless of how they get updated.

    returnedVal = "" & _
        "<menu xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">" & _
            "<menuSeparator id=""ltAutoCorrectMenuBulk"" title=""(!) UPDATE ALL"" />" & _
            "<toggleButton id=""ltAutoCorrectMenuBulkToggle"" getLabel=""ltAutoCorrectMenuBulkToggle_getLabel"" getPressed=""ltAutoCorrectMenuBulkToggle_getPressed"" onAction=""ltAutoCorrectMenuBulkToggle_onAction"" />" & _
            "<menuSeparator id=""ltAutoCorrectMenuGlobalSettings"" title=""(!) GLOBAL SETTINGS"" />" & _
            "<toggleButton id=""ltAutoCorrectMenuReplaceTextToggle"" label=""AutoCorrect: ReplaceText"" getPressed=""ltAutoCorrectMenuReplaceTextToggle_getPressed"" onAction=""ltAutoCorrectReplaceTextToggle_onAction"" />" & _
            "<menuSeparator id=""ltAutoCorrectMenuDocumentSettings"" title=""Document Settings""/>" & _
            "<toggleButton id=""ltAutoCorrectMenuProofingToggle"" label=""Proofing (for 'Normal' Style)"" getPressed=""ltAutoCorrectMenuProofingToggle_getPressed"" onAction=""ltAutoCorrectMenuProofingToggle_onAction"" />" & _
            "<toggleButton id=""ltAutoCorrectMenuShowSpellingErrorsToggle"" label=""ShowSpellingErrors"" getPressed=""ltAutoCorrectMenuShowSpellingErrorsToggle_getPressed"" onAction=""ltAutoCorrectMenuShowSpellingErrorsToggle_onAction"" />" & _
            "<toggleButton id=""ltAutoCorrectMenuShowGrammaticalErrorsToggle"" label=""ShowGrammaticalErrors"" getPressed=""ltAutoCorrectMenuShowGrammaticalErrorsToggle_getPressed"" onAction=""ltAutoCorrectMenuShowGrammaticalErrorsToggle_onAction"" />" & _
        "</menu>"
End Sub

'Check if all auto-correctors we care about are on. If any are off, this returns false.
Private Function ltAutoCorrectMenuBulkToggle_areAllCorrectorsOn()
    'See ltAutoCorrectMenuProofingToggle_getPressed for details of why the NoProofing check is inverted
    ltAutoCorrectMenuBulkToggle_areAllCorrectorsOn = Application.AutoCorrect.ReplaceText _
        And (Not CBool(ActiveDocument.Styles("Normal").NoProofing)) _
        And ActiveDocument.ShowSpellingErrors _
        And ActiveDocument.ShowGrammaticalErrors
End Function

'Callback for ltAutoCorrectMenuBulkToggle getLabel
Sub ltAutoCorrectMenuBulkToggle_getLabel(control As IRibbonControl, ByRef returnedVal)
    returnedVal = "Toggle All " & IIf(ltAutoCorrectMenuBulkToggle_areAllCorrectorsOn, "Off", "On")
End Sub

'Callback for ltAutoCorrectMenuBulkToggle getPressed
Sub ltAutoCorrectMenuBulkToggle_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = ltAutoCorrectMenuBulkToggle_areAllCorrectorsOn
End Sub

'Callback for ltAutoCorrectMenuBulkToggle onAction
Sub ltAutoCorrectMenuBulkToggle_onAction(control As IRibbonControl, pressed As Boolean)
    Application.AutoCorrect.ReplaceText = pressed
    ActiveDocument.Styles("Normal").NoProofing = Not pressed 'See ltAutoCorrectMenuProofingToggle_getPressed for details of why the NoProofing check is inverted
    ActiveDocument.ShowSpellingErrors = pressed
    ActiveDocument.ShowGrammaticalErrors = pressed
End Sub

'Callback for ltAutoCorrectMenuReplaceTextToggle getPressed
Sub ltAutoCorrectMenuReplaceTextToggle_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Application.AutoCorrect.ReplaceText
End Sub

'Callback for ltAutoCorrectMenuReplaceTextToggle onAction
Sub ltAutoCorrectReplaceTextToggle_onAction(control As IRibbonControl, pressed As Boolean)
    Application.AutoCorrect.ReplaceText = pressed
End Sub

'Callback for ltAutoCorrectMenuProofingToggle getPressed
Sub ltAutoCorrectMenuProofingToggle_getPressed(control As IRibbonControl, ByRef returnedVal)
    'NB NoProofing is a Long for some reason, but will only contain bool equivalents. Note also
    '   that we are INVERTING the value, so all the AutoCorrect options for this add-in are positive statements
    '   e.g. Proofing is on, ShowSpellingErrors is on, etc.
    returnedVal = Not CBool(ActiveDocument.Styles("Normal").NoProofing)
End Sub

'Callback for ltAutoCorrectMenuProofingToggle onAction
Sub ltAutoCorrectMenuProofingToggle_onAction(control As IRibbonControl, pressed As Boolean)
    'NB NoProofing is a Long for some reason, but will only contain bool equivalents. Note also
    '   that we are INVERTING the value, so all the AutoCorrect options for this add-in are positive statements
    '   e.g. Proofing is on, ShowSpellingErrors is on, etc.
    ActiveDocument.Styles("Normal").NoProofing = Not pressed
End Sub

'Callback for ltAutoCorrectMenuShowSpellingErrorsToggle getPressed
Sub ltAutoCorrectMenuShowSpellingErrorsToggle_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = ActiveDocument.ShowSpellingErrors
End Sub

'Callback for ltAutoCorrectMenuShowSpellingErrorsToggle onAction
Sub ltAutoCorrectMenuShowSpellingErrorsToggle_onAction(control As IRibbonControl, pressed As Boolean)
    ActiveDocument.ShowSpellingErrors = pressed
End Sub

'Callback for ltAutoCorrectMenuShowGrammaticalErrorsToggle getPressed
Sub ltAutoCorrectMenuShowGrammaticalErrorsToggle_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = ActiveDocument.ShowGrammaticalErrors
End Sub

'Callback for ltAutoCorrectMenuShowGrammaticalErrorsToggle onAction
Sub ltAutoCorrectMenuShowGrammaticalErrorsToggle_onAction(control As IRibbonControl, pressed As Boolean)
    ActiveDocument.ShowGrammaticalErrors = pressed
End Sub
