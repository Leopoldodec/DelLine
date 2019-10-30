Attribute VB_Name = "module_main"
Option Explicit
Global oSkin As Object
Sub AtivaSkin()
    Set oSkin = CreateObject("SkinModuleX.clsSkin")
    oSkin.Licenca = "DJBJ-2801-1904-AAML"
    oSkin.AbrirSkinModo1 (App.Path & "\WinVista.dll")
End Sub

Sub main()
    AtivaSkin
    frm_del_line.Show
End Sub
