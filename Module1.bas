Attribute VB_Name = "Module1"
'Setiap Anda memasukkan suatu karakter di TextBox, maka akan muncul kotak pesan yang menjelaskan tentang jenis 'karakter yang diinput, apakah huruf kecil, huruf 'besar, karakter Alpha, dan Alpha atau Numeric.

Declare Function IsCharUpper Lib "user32" Alias "IsCharUpperA" (ByVal cChar As Byte) As Long

Declare Function IsCharLower Lib "user32" Alias "IsCharLowerA" (ByVal cChar As Byte) As Long

Declare Function IsCharAlpha Lib "user32" Alias "IsCharAlphaA" (ByVal cChar As Byte) As Long

Declare Function IsCharAlphaNumeric Lib "user32" Alias "IsCharAlphaNumericA" (ByVal cChar As Byte) As Long


