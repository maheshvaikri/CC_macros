Attribute VB_Name = "M_R_APET"
Sub APET_najeti()
'2015_05_21

'aby to nebyla epilepticka tabulka
Application.ScreenUpdating = False

'----------------nakopirovat datum a "APET" x-krat pod sebe
ActiveCell.Offset(0, -1).Range("A1:B1").Copy
ActiveCell.Offset(1, -1).Range("A1:B21").Select
ActiveSheet.Paste 'Selection.PasteSpecial Paste:=xlPasteValues by vlozila jen hodnoty, to nechci, pod APET je totiz vzorec

'----------------doplnit sloupec typ
ActiveCell.Offset(-1, 2).Range("A1:A6").Value = "Najeti-voda"
ActiveCell.Offset(5, 2).Range("A1:A2").Value = "Sterilni napoj"
ActiveCell.Offset(7, 2).Range("A1").Value = "Nesterilni sirup"
ActiveCell.Offset(8, 2).Range("A1:A10").Value = "Sterilni lahev"
ActiveCell.Offset(18, 2).Range("A1:A3").Value = "Najeti-vzduch"

'----------------doplnit sloupec komentar
ActiveCell.Offset(-1, 3).Range("A1:A3").Value = "UHT"
ActiveCell.Offset(2, 3).Range("A1:A3").Value = "Rinser"
ActiveCell.Offset(18, 3).Range("A1").Value = "Rinser"
ActiveCell.Offset(19, 3).Range("A1").Value = "UHT"
ActiveCell.Offset(20, 3).Range("A1").Value = "Filler"

'----------------doplnit nuly a NA
ActiveCell.Offset(-1, 4).Range("A1:G22").Value = "N/A"
ActiveCell.Offset(-1, 4).Range("A1:A19").Value = 0
ActiveCell.Offset(-1, 6).Range("A1:B22").Value = 0

'----------mys skoci do patricneho policka
ActiveCell.Offset(5, 3).Range("A1").Select

'---------------vyskoci hlaska na doplneni objemu lahve a typu nápoje/sirupu
MsgBox ("Nezapomen doplnit napoj a sirup a opravit výsledky!")

End Sub


Sub APET_vyjeti()
'2015_05_21

'aby to nebyla epilepticka tabulka
Application.ScreenUpdating = False

'----------------nakopirovat datum a "APET" x-krat pod sebe
ActiveCell.Offset(0, -1).Range("A1:B1").Copy
ActiveCell.Offset(1, -1).Range("A1:B27").Select
ActiveSheet.Paste

'----------------doplnit sloupec typ
ActiveCell.Offset(-1, 0).Range("A1").Select
ActiveCell.Offset(0, 2).Range("A1:A25").Value = "Vyjeti-ster"
ActiveCell.Offset(25, 2).Range("A1:A3").Value = "Vyjeti-vzduch"

'----------------doplnit sloupec komentar
ActiveCell.Offset(0, 3).Range("A1").Value = "Capper 1-5"
ActiveCell.Offset(1, 3).Range("A1").Value = "Capper 6-10"
ActiveCell.Offset(2, 3).Range("A1").Value = "Capper 11-15"
ActiveCell.Offset(3, 3).Range("A1").Value = "Capper 16-20"

ActiveCell.Offset(4, 3).Range("A1").Value = "Filler 1-10"
ActiveCell.Offset(5, 3).Range("A1").Value = "Filler 11-20"
ActiveCell.Offset(6, 3).Range("A1").Value = "Filler 21-30"
ActiveCell.Offset(7, 3).Range("A1").Value = "Filler 31-40"
ActiveCell.Offset(8, 3).Range("A1").Value = "Filler 41-50"
ActiveCell.Offset(9, 3).Range("A1").Value = "Filler 51-60"

ActiveCell.Offset(10, 3).Range("A1").Value = "Rinser 1-10"
ActiveCell.Offset(11, 3).Range("A1").Value = "Rinser 11-20"
ActiveCell.Offset(12, 3).Range("A1").Value = "Rinser 21-30"
ActiveCell.Offset(13, 3).Range("A1").Value = "Rinser 31-40"
ActiveCell.Offset(14, 3).Range("A1").Value = "Rinser 41-50"
ActiveCell.Offset(15, 3).Range("A1").Value = "Rinser 51-60"
ActiveCell.Offset(16, 3).Range("A1").Value = "Rinser 61-70"
ActiveCell.Offset(17, 3).Range("A1").Value = "Rinser 71-80"
ActiveCell.Offset(18, 3).Range("A1").Value = "Rinser 81-90"
ActiveCell.Offset(19, 3).Range("A1").Value = "Rinser 91-100"

ActiveCell.Offset(20, 3).Range("A1").Value = "system dusiku"
ActiveCell.Offset(21, 3).Range("A1").Value = "mrizka"
ActiveCell.Offset(22, 3).Range("A1").Value = "plexisklo"
ActiveCell.Offset(23, 3).Range("A1").Value = "pas za plnicem"
ActiveCell.Offset(24, 3).Range("A1").Value = "predavac uzaveru"

ActiveCell.Offset(25, 3).Range("A1").Value = "Capper"
ActiveCell.Offset(26, 3).Range("A1").Value = "Filler"
ActiveCell.Offset(27, 3).Range("A1").Value = "Rinser"

'----------------doplnit nuly a NA
ActiveCell.Offset(0, 4).Range("A1:B28").Value = "N/A"
ActiveCell.Offset(0, 6).Range("A1:B28").Value = 0
ActiveCell.Offset(0, 8).Range("A1:C28").Value = "N/A"

'---------------vyskoci hlaska na doplneni objemu lahve a typu nápoje/sirupu
MsgBox ("Nezapomen opravit výsledky!")

'--------------posun mys na nasledujici radek
ActiveCell.Offset(28, 0).Range("A1").Select

End Sub

