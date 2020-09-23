Attribute VB_Name = "Systray"
'***************************
'***************************
'IF YOU HAVE ANY QUESTIONS?'
''''''''''EMAIL ME''''''''''
''''JIMIDOTCOM@YAHOO.COM''''
'***************************
'***************************

'user defined type required by Shell_NotifyIcon API call
      Public Type NOTIFYICONDATA
       cbSize As Long
       hwnd As Long
       uId As Long
       uFlags As Long
       uCallBackMessage As Long
       hIcon As Long
       szTip As String * 64
      End Type

      'constants required by Shell_NotifyIcon API call:
      Public Const NIM_ADD = &H0
      Public Const NIM_MODIFY = &H1
      Public Const NIM_DELETE = &H2
      Public Const NIF_MESSAGE = &H1
      Public Const NIF_ICON = &H2
      Public Const NIF_TIP = &H4
      Public Const WM_MOUSEMOVE = &H200
      Public Const WM_LBUTTONDOWN = &H201     'Button down
      Public Const WM_LBUTTONUP = &H202       'Button up
      Public Const WM_LBUTTONDBLCLK = &H203   'Double-click
      Public Const WM_RBUTTONDOWN = &H204     'Button down
      Public Const WM_RBUTTONUP = &H205       'Button up
      Public Const WM_RBUTTONDBLCLK = &H206   'Double-click

      Public Declare Function SetForegroundWindow Lib "user32" _
      (ByVal hwnd As Long) As Long
      Public Declare Function Shell_NotifyIcon Lib "shell32" _
      Alias "Shell_NotifyIconA" _
      (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

      Public nid As NOTIFYICONDATA
      
      '*******************************************************
      '*******************************************************
      '*******************************************************
      'Make the following Property Settings on the same form to which you added the above code:


      'Property         Required Setting for Taskbar Notification Area example
      '-----------------------------------------------------------------------
      'Icon           = The icon you want to appear in the system tray.
      'MinButton = True
      'ShownInTaskbar = False
      '*******************************************************
      '*******************************************************
      '*******************************************************
      
      'Add the following Menu items to the same form using the Menu Editor:


      'Caption      Name          Enabled   Visible   Position
      '---------------------------------------------------------
      '&SysTray     mPopupSys      True      False    Main Level
      '&Restore     mPopRestore    True      True     Inset one
      '&Exit        mPopExit       True      True     Inset one
      'You can add additional menu items as needed.
      '*******************************************************
      '*******************************************************
      '*******************************************************
      'Taskbar Notification Area Flexibility
      'You can modify the ToolTip that appears over the Notification icon by changing the following line in the Form_Load procedure:
      '.szTip = "Your ToolTip" & vbNullChar
      'Replace "Your ToolTip" with the text that you want to appear.

      'You can modify the Icon that appears in the Taskbar Notification Area by changing the following line in the Form_Load procedure:
      '.hIcon = Me.Icon
      'Replace Me.Icon with any Icon in your project.

      'You can change any of the Taskbar Notification Area settings at any time after the use of the NIM_ADD constant by reassigning the values in the nid variable and then using the following variation of the Shell_NotifyIcon API call:
      'Shell_NotifyIcon NIM_MODIFY, nid.
      'However, if you want a different form to receive the callback, then you will need to delete the current icon first using "Shell_NotifyIcon NIM_Delete, nid" as the NIM_Modify function will not accept a new Hwnd, or you will need to add another Icon to the systray for the new form using "Shell_NotifyIcon NIM_ADD, nid" after refilling the nid type with the new forms Hwnd. You can also declare separate copies of the nid type for each form that you want to display an icon for in the Windows System Tray and change them in each form's activate event using the NIM_DELETE and NIM_ADD sequence.






