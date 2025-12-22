
<img width="790" height="937" alt="image" src="https://github.com/user-attachments/assets/32d3d201-8052-480a-91ed-c6e31458ba23" />

帮助文档：
C:\Users\YLIU119\OneDrive - azureford\desktop\Catia tool\online

## 创建catiavba 文档
<img width="640" height="300" alt="image" src="https://github.com/user-attachments/assets/241d6c86-b256-40ad-9c1b-a0b2dffaf52d" />


### 第一个测试文件，获取文件名：

    Sub GetCatiaChild()
    
    Dim oproductDoc As Document
    Set oproductDoc = CATIA.ActiveDocument ' returns a Document
    
    Set oProduct = oproductDoc.Product
    
    Dim TopProductName As String
    TopProductPN = oProduct.PartNumber
    
    MsgBox TopProductPN
    
    If oProduct.Products.Count >= 1 Then
        Set oSubProduct = oProduct.Products.Item(1)
        MsgBox oSubProduct.PartNumber
    End If
    
    End Sub
