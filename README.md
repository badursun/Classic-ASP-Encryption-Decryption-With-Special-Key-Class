# Classic ASP Encryption &amp; Decryption With Special Key Class

## Introduction
RES (Rabbit Encryption System) is a specially compiled encryption class. It was developed through an anonymous algorithm similar to the RC4 algorithm. It provides easier handling in a tidy class.

You are free to develop as you wish. It provides you with all the parameters you need with just a few small codes.

For advanced use; You can create a key file for each defined user or store it in the database of that user. Thus, each encrypted data becomes unique.

I recommend using a password of at least 512 characters.

## Usage 
Import class your project
```asp
<!--include file="/yourpath/res.asp"-->
```

Set the variable
```asp
Set Encryption = New RabbitEncryptionSystem
```

And use it as you wish to use.

## Example
```asp
<%
  Set Encryption = New RabbitEncryptionSystem
    SpecialWords    = "Bu gizli bir kelimedir, gizli olarak kalması gerekmektedir!"

    ' If you want create random Key uncomment end define empty string
    ' Or Define your Key. More Length, More Strength (512 Length)
    ' Encryption.Key  = ""
    Encryption.Key  = "/[g-?#4Fd$-T/{\d3%%.@!specialLongKey$%&₺sW"

    Response.Write "<strong>Default Key:</strong> " & Encryption.DefaultKey()
    Response.Write "<hr>"

    Response.Write "<strong>Your Key:</strong> " & Encryption.Key()
    Response.Write "<hr>"
    Response.Write "<strong>Special Words:</strong> " & SpecialWords
    Response.Write "<hr>"
    Response.Write "<strong>Encrypted:</strong> " & Encryption.Encrypt(SpecialWords)
    Response.Write "<hr>"
    Response.Write "<strong>Decrypted:</strong> " & Encryption.Decrypt( Encryption.Encrypt(SpecialWords) )
    Response.Write "<hr>"

    ' If you want save key from file, define your top-secret-folder
    'Encryption.FilePath = "../cache/"
    Response.Write "Key Save Path: " & Encryption.KeySavePath()
    Response.Write "<hr>"

    ' Save File To Path & File Type
    MyKeyFileName = "mykey.txt"
    
    ' Write File
    Encryption.WriteKeyToFile(MyKeyFileName)
    
    Response.Write "<strong>File Saved:</strong> " & Encryption.FileSaved()
    Response.Write "<hr>"
    Response.Write "<strong>Last Saved File Name:</strong> " & Encryption.LastSavedFile()
    Response.Write "<hr>"

    ' Read File For Key
    Response.Write "<strong>Readed Key File:</strong> " & MyKeyFileName
    Response.Write "<hr>"
    Response.Write "<strong>Readed Key Value:</strong> " & Encryption.ReadKeyFromFile(MyKeyFileName)
    Response.Write "<hr>"
    Response.Write "<strong>Not Exist Readed Key File:</strong> " & Encryption.ReadKeyFromFile("NotExistFile.txt")
    Response.Write "<hr>"

  Set Encryption = Nothing
%>
```

## Output Example

```txt
Default Key: GNQ?4i0-*\CldnU+[vrF1j1PcWeJfVv4QGBurFK6}*l[H1S:oY\v@U?i,oD]f/n8oFk6NesH--^PJeCLdp+(t8SVe:ewY(wR9p-CzG<,Q/(U*.pXDiz/KvnXP`BXnkgfeycb)1A4XKAa-2G}74Z8CqZ*A0P8E[S`6RfLwW+Pc}13U}_y0bfscJW/m1D*YoAlkBK@`3A)trZsO5xv@5@MRRFkt\

Key: /[g-?#4Fd$-T/{\d3%%.@!specialLongKey$%&₺sW

Keyword: Bu gizli bir kelimedir, gizli olarak kalması gerekmektedir!

Encrypted: A%A0Wdxmp%7FTVf%96%1F%B6%91%A0lbZbyco%60%9C%9C%B3%9D%A5%3C%AE%AA%98%8D%96%B4%14%60W%1F%BF%EB%B6%EB%B0K%9Eb%81Xo%83%99%5Fq%89c%B4%9EU

Decrypted: Bu gizli bir kelimedir, gizli olarak kalması gerekmektedir!

Key Save Path: C:\Inetpub\vhosts\mydomain.com\httpdocs\content\keyFiles

File Saved: True

Last Saved File Name: mykey.txt

Readed Key File: mykey.txt

Readed Key Value: /[g-?#4Fd$-T/{\d3%%.@!specialLongKey$%&₺sW

Not Exist Readed Key File: File Not Found
```
