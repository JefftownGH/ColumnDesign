/*namespace ColumnDesign.Modules
{
    public class KeypressModuleFtIn
    {
        Public Function KeyInputRestrictorFtIn(ByVal KeyAscii As MSForms.ReturnInteger, BoxValue)
        If KeyAscii < 48 Or KeyAscii > 57 Then                                                              'If the key is a number it is allowed. If the key is not a number then...
        If BoxValue = "" And (KeyAscii = 34 Or KeyAscii = 39 Or KeyAscii = 45 Or KeyAscii = 46) Then        'If this keypress is the first character and it's a ' " - or . then...
        KeyAscii = 0                                                                                        'BLOCK the keypress
            End If
            If KeyAscii = 34 Or KeyAscii = 39 Or KeyAscii = 45 Or KeyAscii = 46 Then                            'If key is " ' - or . then ALLOW the keypress
        Else                                                                                                'If key is not number and not " ' - or . then...
        KeyAscii = 0                                                                                        'BLOCK the keypress
            End If
            End If
            End Function

    }
}*/