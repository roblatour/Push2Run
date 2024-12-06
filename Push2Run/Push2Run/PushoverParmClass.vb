Public Class PushoverIdAndSecretParmClass

    Public email As String
    Public password As String
    Public twoFactorAuthenticationCode As String

End Class

Public Enum ForceValues
    NoForce = 0
    Force = 1
    ForceApplied = 2
End Enum

Public Class PushoverDeviceNameClass

    Public name As String
    Public force As ForceValues

End Class




