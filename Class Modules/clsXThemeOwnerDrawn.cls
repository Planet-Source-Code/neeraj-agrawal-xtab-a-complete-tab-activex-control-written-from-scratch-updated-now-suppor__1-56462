VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "xThemeOwnerDrawn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'xThemeOwnerDrawn : Calls events to allow client to draw the complete control
'                   But NOT IMPLEMENTED COMPLETELY. Since Contained controls need to be handled
'                   See ITheme for implemented function description

Option Explicit

Implements ITheme

Private m_oXTab As XTab

'===Implemented Functions=============================================================================================

Private Sub ITheme_DrawBackground()
  m_oXTab.bAutoRedraw = False
  m_oXTab.pRaiseDrawBackground
End Sub

Private Sub ITheme_DrawBackgroundPropertyPages()
  'no need in the Owner Drawn Theme
End Sub

Private Sub ITheme_DrawBackgroundTabbedDialog()

End Sub

'these functions were seperated acc to theme so as to allow painting of only the desired parts (and not the whole control)
Private Sub ITheme_DrawOnActiveTabChange()
  m_oXTab.bAutoRedraw = False
  m_oXTab.pRaiseDrawOnActiveTabChange
  'm_oXTab.bAutoRedraw = True
End Sub

Private Sub ITheme_DrawTabs()
  m_oXTab.bAutoRedraw = False
  m_oXTab.pRaiseDrawTabs
  'm_oXTab.bAutoRedraw = True
End Sub

Private Sub Itheme_DrawTabsPropertyPages()

End Sub

Private Sub ITheme_DrawTabsTabbedDialog()

End Sub

Private Sub ITheme_MouseDownHanlder(iButton As Integer, iShift As Integer, sngX As Single, sngY As Single)

End Sub

Private Sub ITheme_MouseMoveHanlder(iButton As Integer, iShift As Integer, sngX As Single, sngY As Single)

End Sub

Private Sub ITheme_MouseUpHanlder(iButton As Integer, iShift As Integer, sngX As Single, sngY As Single)
  'Do Nothing
End Sub

Private Sub ITheme_SetControl(oXTab As XTab)
  Dim utRect As RECT
  
  Set m_oXTab = oXTab
  
  'must set the following properties for the owner drawn control
  m_oXTab.bAutoRedraw = True
  m_oXTab.pCls                  'clear any existing things
  m_oXTab.bAutoRedraw = False
  
  If Not m_oXTab.bUserMode Then
    m_oXTab.bAutoRedraw = True
    utRect.Right = m_oXTab.iScaleWidth
    utRect.Bottom = m_oXTab.iScaleHeight
    
    m_oXTab.pLine 0, 0, m_oXTab.iScaleWidth - 1, m_oXTab.iScaleHeight - 1, m_oXTab.OuterBorderColor, True, False
    
    DrawText m_oXTab.lHDC, "Owner Drawn not available at design time", -1, utRect, DT_VCENTER Or DT_SINGLELINE Or DT_CENTER Or DT_END_ELLIPSIS
  End If
  
  
End Sub

Private Sub ITheme_ShowHideFocus()
  'm_oXTab.bAutoRedraw = False
  m_oXTab.pRaiseShowHideFocus
  'm_oXTab.bAutoRedraw = True
End Sub

Private Sub ITheme_ShowHideFocusPropertyPages()

End Sub

Private Sub ITheme_ShowHideFocusTabbedDialog()

End Sub


Private Sub ITheme_TimerEvent()

End Sub


Private Sub ITheme_ResetColorsToDefault()
  'do nothing
End Sub
'=====================================================================================================================



