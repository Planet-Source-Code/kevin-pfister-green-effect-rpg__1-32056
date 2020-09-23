Attribute VB_Name = "ModGreen"
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Public Type WType
    Name As String
    Attack As Double
    max As Double
    Used As Double
    MissChance As Double
End Type

Public Type AType
    Name As String
    Defence As Double
    max As Double
    Used As Double
End Type

Dim Weapon(1 To 10) As WType
Dim Armour As AType

Public PlayerHealth    'The Players Health
Public PlayerArmour    'The Players Armour
Public PlayerWeapon    'The Uses left of weapon
Public Castras          'Money

Sub BuyArmour(index)
    If index = 1 Then 'Cloth armour
        Armour.Defence = 2
        Armour.max = 10
        Armour.Used = 0
        
        Castras = Castras - 10
    ElseIf index = 2 Then   'Leather armour
        Armour.Defence = 5
        Armour.max = 20
        Armour.Used = 0
        
        Castras = Castras - 25
    ElseIf index = 3 Then 'chainmail
        Armour.Defence = 10
        Armour.max = 40
        Armour.Used = 0
        
        Castras = Castras - 100
    End If
    PlayerArmour = Armour.max - Armour.Used
    Call updateDraw
    Call DoSell
End Sub

Sub buyweapon(index)

End Sub

Sub updateDraw()
    If FrmGreenEffect.ProgressFore(1).Width <> Int((FrmGreenEffect.ProgressBack(1).Width / 100) * PlayerArmour) Then
        FrmGreenEffect.ProgressFore(1).Width = Int((FrmGreenEffect.ProgressBack(1).Width / 100) * PlayerArmour)
    End If
    If FrmGreenEffect.ProgressFore(2).Width <> Int((FrmGreenEffect.ProgressBack(2).Width / 100) * PlayerWeapon) Then
        FrmGreenEffect.ProgressFore(2).Width = Int((FrmGreenEffect.ProgressBack(2).Width / 100) * PlayerWeapon)
    End If
    If FrmGreenEffect.ProgressFore(3).Width <> Int((FrmGreenEffect.ProgressBack(3).Width / 100) * PlayerHealth) Then
        FrmGreenEffect.ProgressFore(3).Width = Int((FrmGreenEffect.ProgressBack(3).Width / 100) * PlayerHealth)
    End If
    If FrmGreenEffect.lblMoney.Caption <> Castras Then
        FrmGreenEffect.lblMoney.Caption = Castras
    End If
End Sub

Sub Sell()

End Sub

Sub NoSell()
    FrmGreenEffect.LblMessage.Caption = "Talking to Shop Keeper"
    FrmGreenEffect.LblText.Caption = "Sorry you can't afford that, why don't you come back later on..."
    TimeBefore = Timer
    While Timer < TimeBefore + 2.5
        DoEvents
    Wend
    FrmGreenEffect.LblMessage.Caption = ""
    FrmGreenEffect.LblText.Caption = ""
End Sub

Sub DoSell()
    FrmGreenEffect.LblMessage.Caption = "Talking to Shop Keeper"
    FrmGreenEffect.LblText.Caption = "Thank You... Would you be interested in anything else?"
    TimeBefore = Timer
    While Timer < TimeBefore + 2.5
        DoEvents
    Wend
    FrmGreenEffect.LblMessage.Caption = ""
    FrmGreenEffect.LblText.Caption = ""
End Sub

Sub DoBCase()
    FrmGreenEffect.LblMessage.Caption = "Looking at a BookCase"
    FrmGreenEffect.LblText.Caption = "Such an interesting Bookcase, shame i can't read the writing"
    TimeBefore = Timer
    While Timer < TimeBefore + 2.5
        DoEvents
    Wend
    FrmGreenEffect.LblMessage.Caption = ""
    FrmGreenEffect.LblText.Caption = ""
End Sub

Sub DoCase()
    FrmGreenEffect.LblMessage.Caption = "Looking at a chest"
    FrmGreenEffect.LblText.Caption = "Lots of different objects, wish i had some"
    TimeBefore = Timer
    While Timer < TimeBefore + 2.5
        DoEvents
    Wend
    FrmGreenEffect.LblMessage.Caption = ""
    FrmGreenEffect.LblText.Caption = ""
End Sub

Sub DoBed()
    FrmGreenEffect.LblMessage.Caption = "Looking at the Bed"
    FrmGreenEffect.LblText.Caption = "Such a comfortable Bed, but i don't have time to rest"
    TimeBefore = Timer
    While Timer < TimeBefore + 2.5
        DoEvents
    Wend
    FrmGreenEffect.LblMessage.Caption = ""
    FrmGreenEffect.LblText.Caption = ""
End Sub

