<img width="565" height="428" alt="image" src="https://github.com/user-attachments/assets/5b1b78fc-aaab-40bb-be89-246b4ff598fe" />


' ========== 全局变量（仅定义1次） ==========
Public g_topProduct As Product
Public g_copiedProduct As Product
Public g_ProductDoc As ProductDocument
Public g_copiedPart As Part
Public oSel As Selection    'In Proudct, part to part COpy, must use top productdoc selection
Public g_stop As Boolean
 

Private Sub CommandButton1_Click()
    ' 1. 创建新的顶级产品文档
    Dim documents1 As Documents
    Set documents1 = CATIA.Documents
    
    Dim selectedFilePath As String
    selectedFilePath = CATIA.FileSelectionBox("Plese Select CATProduct/CATPart", "*.CATProduct;*.CATPart;*.*", CatFileSelectionModeOpen)
    If selectedFilePath = "" Then
        MsgBox "你取消了文件选择，程序终止", vbInformation
        g_stop = True
        
        Exit Sub
    End If

    g_stop = False
    Dim productDocument1 As ProductDocument
    
    
    
    Dim existingdoc   'product or part doc
'   Dim existingdoc1 As ProcessDocument
'   existingdoc1.Activate
    
    Dim product1 As Product
    
    Dim i As Integer
    Dim flg As Boolean
    flg = False

    For i = 1 To documents1.Count
         
        Set existingdoc = documents1.Item(i)
        If existingdoc.Name = "TopAssemblyForExtractASurface.CATProduct" Then
            Set product1 = existingdoc.Product
            existingdoc.Activate
            flg = True
            Exit For
        End If
        
    Next i

    If flg = False Then
    
           Set productDocument1 = documents1.Add("Product")
           Set product1 = productDocument1.Product
           product1.PartNumber = "TopAssemblyForExtractASurface"
    End If

    Set g_topProduct = product1
    Set g_ProductDoc = g_topProduct.ReferenceProduct.Parent
    g_topProduct.ApplyWorkMode DESIGN_MODE

    ' 4. 创建copiedPart（替代原PartForCopy，作为A面提取目标）
    If Not CreateCopiedPart() Then
        MsgBox "copiedPart创建失败，程序终止", vbCritical
        Exit Sub
    End If

 
'    ' 2. 获取顶级产品的子组件集合并导入外部CATProduct文件
    Dim products1 As products
    Set products1 = product1.products
    Dim arrayOfVariantOfBSTR1(0) As Variant
''    arrayOfVariantOfBSTR1(0) = "C:\Users\YLIU119\OneDrive - azureford\desktop\SeatCAD\RU5A-618A13-C.CATProduct"
    arrayOfVariantOfBSTR1(0) = selectedFilePath
    Set products1Variant = products1
    products1Variant.AddComponentsFromFiles arrayOfVariantOfBSTR1, "All"

    Call updateListBox1
    
  

    ' 3. 初始化全局顶层装配变量（对接原A面提取逻辑）
'    Set g_topProduct = product1
'    Set g_ProductDoc = g_topProduct.ReferenceProduct.Parent
'    g_topProduct.ApplyWorkMode DESIGN_MODE
'
'    ' 4. 创建copiedPart（替代原PartForCopy，作为A面提取目标）
'    If Not CreateCopiedPart() Then
'        MsgBox "copiedPart创建失败，程序终止", vbCritical
'        Exit Sub
'    End If
    
'     ListBox1.AddItem (products1Variant.Item(2).Name)
    

End Sub


' ========== 整合函数：创建copiedPart+目标几何集 ==========
Private Function CreateCopiedPart() As Boolean
    ' 默认返回失败
    CreateCopiedPart = False
    
    Dim i As Integer
    Dim existingpartdoc
    Dim flg2 As Boolean
    flg2 = False
    
    For i = 1 To g_topProduct.products.Count
        Set existingpartdoc = g_topProduct.products.Item(i)
        If existingpartdoc.Name = "PartforCopiedSurface" Then
         Set g_copiedProduct = existingpartdoc '.Products
             ' 赋值全局g_copiedPart
         Set g_copiedPart = g_copiedProduct.ReferenceProduct.Parent.Part
         flg2 = True
         Exit For
        End If
    Next i
    
    ' 创建copiedPart组件（添加到顶级装配）
    If flg2 = False Then
        Set g_copiedProduct = g_topProduct.products.AddNewComponent("Part", "")
        g_copiedProduct.Name = "PartforCopiedSurface"
        g_copiedProduct.PartNumber = "Extracted_A_Faces_Part"
        Set g_copiedPart = g_copiedProduct.ReferenceProduct.Parent.Part
                
        ' 创建A面存储的目标几何集（替代原PartForCopy里的CopyASurface）
        Dim hb As HybridBody
        Set hb = g_copiedPart.HybridBodies.Add()
        hb.Name = "All_Extracted_A_Faces"
    End If

    CreateCopiedPart = True
End Function


Private Sub CommandButton2_Click()
 
    Dim i As Long
    Dim key As String

    ' === 先构建字典，装入 ListBox2 现有项，用于重复检查 ===
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    For i = 0 To ListBox2.ListCount - 1
        dict(ListBox2.List(i, 0)) = True   ' 单列显式用 (i,0)
    Next i

    ' === 发送选中项到 ListBox2（仅当不重复时添加） ===
    For i = 0 To ListBox1.ListCount - 1
        If ListBox1.Selected(i) Then
            key = ListBox1.List(i, 0)      ' 第一列内容
            If Not dict.Exists(key) Then
                ListBox2.AddItem key
                dict(key) = True           ' 同批次更新，避免本次重复
            End If
        End If
    Next i

    ' === 发送后取消选择（保持你的过程内完成，不调用外部方法） ===
    For i = 0 To ListBox1.ListCount - 1
        If ListBox1.Selected(i) Then
            ListBox1.Selected(i) = False
        End If
    Next i
    ListBox1.ListIndex = -1
    
 ' 同步禁用按钮（因为现在没有选中项了）
    Me.CommandButton2.Enabled = False


    ' === 用（已更新过的）字典标记 ListBox1 的第二列状态 ===
    Dim partNumber_List1 As String
    For i = 0 To ListBox1.ListCount - 1
        partNumber_List1 = ListBox1.List(i, 0)
        If dict.Exists(partNumber_List1) Then
            ListBox1.List(i, 1) = "√"
        Else
            ListBox1.List(i, 1) = ""
        End If
    Next i

End Sub


Private Sub CommandButton3_Click()


Dim i As Long

    ' 1) 倒序删除 ListBox2 选中的项
    For i = Me.ListBox2.ListCount - 1 To 0 Step -1
        If Me.ListBox2.Selected(i) Then
            Me.ListBox2.RemoveItem i
            Exit For
        End If
    Next i



    ' 2) 用 ListBox2 的当前内容构建字典
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")

    If Me.ListBox2.ListCount > 0 Then
        For i = 0 To Me.ListBox2.ListCount - 1
            ' 显式第一列：避免多列时 .List(i) 取不到值
            dict(Me.ListBox2.List(i, 0)) = True
        Next i
    End If

    ' 3) 遍历 ListBox1，按字典标记重复
    Dim partNumber_List1 As String
    For i = 0 To Me.ListBox1.ListCount - 1
        partNumber_List1 = Me.ListBox1.List(i, 0)  ' 第一列（件号）
        If dict.Exists(partNumber_List1) Then
            Me.ListBox1.List(i, 1) = "√"           ' 第二列显示状态
        Else
            Me.ListBox1.List(i, 1) = ""            ' 清空状态
        End If
    Next i
    
    
        ' === 发送后取消选择（保持你的过程内完成，不调用外部方法） ===
    For i = 0 To ListBox2.ListCount - 1
        If ListBox2.Selected(i) Then
            ListBox2.Selected(i) = False
        End If
    Next i

    Me.CommandButton3.Enabled = (Me.ListBox2.ListIndex >= 0 And Me.ListBox2.ListCount > 0)

End Sub

'show loaded parts

Private Sub CommandButton4_Click()
    Call updateListBox1
    
End Sub


Private Sub updateListBox1()
    Dim windows1 As Windows
    Set windows1 = CATIA.Windows
    Dim w_count As Long
    w_count = windows1.Count
    
    Dim specsAndGeomWindow1 As SpecsAndGeomWindow
    Dim productDocument1 As ProductDocument
    Dim topProdcut As Product
    Dim childProduct As Product
'    Dim products As products
    
    Dim arrayOfVariantOfBSTR2(0)
  
 

    Dim i As Integer
    Dim j As Integer
    Dim PartNumber As String
    ListBox1.Clear
    For i = 1 To w_count
        Set specsAndGeomWindow1 = windows1.Item(i)
        If specsAndGeomWindow1.Name = "TopAssemblyForExtractASurface.CATProduct" Then
            Set productDocument1 = specsAndGeomWindow1.Parent
'
            Set topProduct = productDocument1.Product

                For Each childProduct In topProduct.products
                    PartNumber = childProduct.PartNumber
                    If childProduct.PartNumber <> "Extracted_A_Faces_Part" Then
                        ListBox1.AddItem (PartNumber)

                    ' 刚添加的项在最后一行，索引 = ListCount - 1
                        idx = Me.ListBox1.ListCount - 1
                        ListBox1.List(idx, 1) = " "   'key is idx, status
                    End If
                Next
            Exit For

        End If
    Next i
     
End Sub


Private Sub CommandButton5_Click()
    ListBox2.Clear
    
        ' === 用（已更新过的）字典标记 ListBox1 的第二列状态 ===
    Dim partNumber_List1 As String
    For i = 0 To ListBox1.ListCount - 1
           ListBox1.List(i, 1) = ""

    Next i

End Sub



' UserForm 初始化时设置 ListBox1 为两列并填充数据
Private Sub UserForm_Initialize()
    ' === 设置为两列，并设置每列宽度 ===
    With Me.ListBox1
        .ColumnCount = 2
        ' 宽度可以用 pt（点）或 twips；pt更直观。根据你的窗体宽度调整数值
        .ColumnWidths = "120 pt;20 pt"
        .MultiSelect = fmMultiSelectExtended
        .Clear
    End With

    Me.ListBox1.ListIndex = -1
    
    ' 初始化时禁用按钮（没有选中项）
        Me.CommandButton2.Enabled = False
        Me.CommandButton3.Enabled = False



End Sub


Private Sub ListBox1_Change()
    Dim hasSelection As Boolean
    Dim i As Long

    If Me.ListBox1.MultiSelect = fmMultiSelectSingle Then
        ' 单选：有选中项则启用
        hasSelection = (Me.ListBox1.ListIndex >= 0)
    Else
        ' 多选/扩展：遍历 Selected(i)
        hasSelection = False
        For i = 0 To Me.ListBox1.ListCount - 1
            If Me.ListBox1.Selected(i) Then
                hasSelection = True
                Exit For
            End If
        Next i
    End If

    Me.CommandButton2.Enabled = hasSelection
End Sub

Private Sub ListBox2_Change()
    Me.CommandButton3.Enabled = (Me.ListBox2.ListIndex >= 0 And Me.ListBox2.ListCount > 0)
End Sub




Private Sub ExtractAllCAD_Click()
       
    ' 1. 创建新的顶级产品文档
    Dim documents1 As Documents
    Set documents1 = CATIA.Documents
    
    Dim productDocument1 As ProductDocument
    
    
    Dim existingdoc
    
    Dim product1 As Product
    Dim products1 As products
    
    Dim i As Integer
    Dim flg As Boolean
    flg = False

    For i = 1 To documents1.Count
         
        Set existingdoc = documents1.Item(i)
        If existingdoc.Name = "TopAssemblyForExtractASurface.CATProduct" Then
            Set product1 = existingdoc.Product
            existingdoc.Activate
            flg = True
            Exit For
        End If
        
        
    Next i

    If flg = False Then
    
            Call CommandButton1_Click
            Set product1 = g_topProduct
            If g_stop = True Then Exit Sub
        
    End If

 

'    ' 2. 获取顶级产品的子组件集合并导入外部CATProduct文件
    
    Set products1 = product1.products
  

    ' 3. 初始化全局顶层装配变量（对接原A面提取逻辑）
    Set g_topProduct = product1
    Set g_ProductDoc = g_topProduct.ReferenceProduct.Parent
    g_topProduct.ApplyWorkMode DESIGN_MODE

    ' 4. 创建copiedPart（替代原PartForCopy，作为A面提取目标）
    If Not CreateCopiedPart() Then
        MsgBox "copiedPart创建失败，程序终止", vbCritical
        Exit Sub
    End If

    ' 5. 递归遍历装配并提取所有零件的A面到copiedPart
    Call TraverseAndExtractFaces(g_topProduct, True, "")

    ' 6. 最终更新并提示完成
    g_copiedPart.update
    MsgBox "装配加载完成！A面已全部提取到copiedPart的All_Extracted_A_Faces几何集", vbInformation
End Sub




Private Sub ExtractSelectedCAD_Click()

    Dim partList() As String ' 定义字符串数组（存储Name）
    ReDim partList(0 To Me.ListBox2.ListCount - 1)
    Dim i As Integer
    For i = 0 To Me.ListBox2.ListCount - 1
        partList(i) = CStr(Me.ListBox2.List(i, 0))
        partList(i) = Trim(partList(i))
    Next
    
    ' 1. 创建新的顶级产品文档
    Dim documents1 As Documents
    Set documents1 = CATIA.Documents
    
    Dim productDocument1 As ProductDocument
    
    Dim existingdoc
    
    Dim product1 As Product
    Dim products1 As products
    
  
    Dim flg As Boolean
    flg = False

    For i = 1 To documents1.Count
         
        Set existingdoc = documents1.Item(i)
        If existingdoc.Name = "TopAssemblyForExtractASurface.CATProduct" Then
            Set product1 = existingdoc.Product
            existingdoc.Activate
            flg = True
            Exit For
        End If
    Next i

    If flg = False Then
            Call CommandButton1_Click
            Set product1 = g_topProduct
            If g_stop = True Then Exit Sub
    End If

 

    
    Set products1 = product1.products
  

    ' 3. 初始化全局顶层装配变量（对接原A面提取逻辑）
    Set g_topProduct = product1
    Set g_ProductDoc = g_topProduct.ReferenceProduct.Parent
    g_topProduct.ApplyWorkMode DESIGN_MODE

    ' 4. 创建copiedPart（替代原PartForCopy，作为A面提取目标）
    If Not CreateCopiedPart() Then
        MsgBox "copiedPart创建失败，程序终止", vbCritical
        Exit Sub
    End If

    ' 5. 递归遍历装配并提取所有零件的A面到copiedPart
    Call TraverseAndExtractFaces(g_topProduct, False, partList)

    ' 6. 最终更新并提示完成
    g_copiedPart.update
    MsgBox "A面已全部提取到copiedPart！", vbInformation
End Sub


' ========== 提取A面核心函数（供ExtractAllCAD_Click调用） ==========
Public Sub ExtractAFaces()
    If g_topProduct Is Nothing Then
        MsgBox "请先加载装配文件", vbExclamation
        Exit Sub
    End If
    If g_copiedPart Is Nothing Then
        MsgBox "目标零件未创建", vbExclamation
        Exit Sub
    End If

    ' 递归遍历选中的零件（或全部零件）并提取A面
    Call TraverseAndExtractFaces(g_topProduct)

    ' 更新并提示
    g_copiedPart.update
    MsgBox "A面提取完成！所有面已保存到PartforCopiedSurface", vbInformation
End Sub





' ========== 核心整合函数：递归遍历+处理零件+提取A面 ==========
Private Sub TraverseAndExtractFaces(ByVal currentProduct As Product, Optional ByVal isALL As Boolean = True, _
    Optional ByVal partLists As Variant = Empty, Optional ByVal targetPartName As String = "")

    On Error GoTo NextProduct
    If currentProduct Is Nothing Then Exit Sub
    
        ' 【新增】初始化字典：缓存需要提取的零件编号（提升匹配效率） ' 仅当 isALL=False 且 partLists 非空时，才创建字典

    Dim partDict As Object
    If Not isALL And Not IsEmpty(partLists) Then
        Set partDict = CreateObject("Scripting.Dictionary")
        ' 遍历数组，将零件编号存入字典（键=零件号，值=True）
        Dim partNum As Variant
        For Each partNum In partLists
            If Not partDict.Exists(CStr(partNum)) Then
                partDict(CStr(partNum)) = True
            End If
        Next
    End If
    
    Dim childProduct As Product
    For Each childProduct In currentProduct.products
    
        Dim isProcess As Boolean
        isProcess = True ' 默认处理
        
        targetPartName = childProduct.PartNumber
        
        If Not isALL Then
            ' isALL=False时，仅匹配partLists中的零件编号
            isProcess = False
            If Not partDict Is Nothing And partDict.Exists(childProduct.PartNumber) Then
                isProcess = True
            End If
        End If
        
        ' 仅当需要处理时，才执行后续逻辑
        If isProcess Then
        
            Dim currentLevelPartNumber As String
        
            ' 子产品仍有下级装配，继续递归
            If childProduct.products.Count > 0 Then
                Call TraverseAndExtractFaces(childProduct, True, "", "")
            Else
                ' 叶子节点零件，直接处理A面提取
                If childProduct.ReferenceProduct Is Nothing Then GoTo NextProduct
                If Not TypeName(childProduct.ReferenceProduct.Parent) = "PartDocument" Then GoTo NextProduct
                
                Dim srcPartDoc As PartDocument
                Set srcPartDoc = childProduct.ReferenceProduct.Parent
                srcPartDoc.Activate ' 激活源Part文档
                
                ' 提取当前零件的A面
                Call ExtractAFacesFromPart(srcPartDoc, targetPartName)
        End If
            
NextProduct:
        End If
    Next
    ' 释放字典对象
    Set partDict = Nothing
End Sub

' ========== 整合函数：提取单个零件的所有A面 ==========
Private Sub ExtractAFacesFromPart(partDoc As PartDocument, childProductPartName As String)
    Dim partObj As Part: Set partObj = partDoc.Part
    Dim body As body
    Dim sel As Selection: Set sel = partDoc.Selection
    
    
    Dim visProps As VisPropertySet
    
    Dim bodyRef As Reference
 
    Dim showstate As CatVisPropertyShow

    
    For Each body In partObj.bodies
        
        Set bodyRef = partDoc.Part.CreateReferenceFromObject(body)
         
        If Not body.InBooleanOperation Then
            ' 清空选择集，添加当前body并搜索面
            sel.Clear: sel.Add body
            sel.VisProperties.GetShow showstate ' catVisPropertyShowAttr
            If showstate = 1 Then GoTo NextBody  'showstate=0 显示，1为隐藏
            
            sel.Search "Topology.Face,sel"
            If sel.Count2 = 0 Then GoTo NextBody
            
            ' 查找有效种子面
            Dim refSeed As Reference, t As Integer
            For t = sel.Count2 To 1 Step -1
                Set refSeed = sel.Item2(t).Reference
                If TypeName(refSeed.Parent) <> "AxisSystem" And IsMeasurableFace(partDoc, refSeed) Then Exit For
            Next
            If refSeed Is Nothing Or t = 0 Then GoTo NextBody
            
            ' 创建Extract曲面
            Dim hsf As HybridShapeFactory: Set hsf = partObj.HybridShapeFactory
            Dim ext As HybridShapeExtract: Set ext = hsf.AddNewExtract(refSeed)
            ext.PropagationType = 1: ext.IsFederated = True
            
            ' 创建临时几何集并添加Extract曲面
            Dim hb As HybridBody: Set hb = partObj.HybridBodies.Add()
            hb.Name = "Temp_Extract_" & body.Name
            hb.AppendHybridShape ext: partObj.UpdateObject ext
            
            ' 创建显式曲面
            Dim explicitSurf As HybridShapeSurfaceExplicit
            Set explicitSurf = hsf.AddNewSurfaceDatum(partObj.CreateReferenceFromObject(ext))
            explicitSurf.Name = "Explicit_" & body.Name
            hb.AppendHybridShape explicitSurf: partObj.UpdateObject explicitSurf
            
            ' 复制曲面到copiedPart
            Call CopySurfaceToTarget(explicitSurf, partObj, partDoc, childProductPartName)
            
            ' 清理临时对象
            g_copiedPart.update
            hsf.DeleteObjectForDatum ext
            sel.Clear
            
NextBody:
        End If
    Next
End Sub

' ========== 整合函数：复制曲面到copiedPart ==========
Private Sub CopySurfaceToTarget(srcSurf As HybridShapeSurfaceExplicit, srcPart As Part, srcPartDoc As PartDocument, Optional ByVal targetPartName As String = "")
    ' 基础校验
    If g_copiedPart Is Nothing Then
        Debug.Print "目标copiedPart未初始化"
        Exit Sub
    End If
    If srcSurf Is Nothing Then
        Debug.Print "源曲面为空"
        Exit Sub
    End If
    
        ' 校验目标名称（为空则用默认名称）
    If Trim(targetPartName) = "" Then
        targetPartName = "All_Extracted_A_Faces"
    Else
        ' 清理非法字符（CATIA几何集名称不支持的字符）
        targetPartName = CleanInvalidChars(targetPartName)
    End If
    
    ' 执行复制粘贴操作
'    CATIA.StartCommand "Fit All In"
    Set oSel = g_ProductDoc.Selection
    oSel.Clear
    oSel.Add srcSurf
    oSel.Copy
    oSel.Clear
    
    ' 粘贴到copiedPart的目标几何集
'    oSel.Add g_copiedPart.HybridBodies.Item("All_Extracted_A_Faces")
'     oSel.Add g_copiedPart.HybridBodies.Item(targetPartName)
'    oSel.PasteSpecial "CATPrtResultWithOutLink"
'    CATIA.RefreshDisplay = True

    ' ======================
       ' 核心逻辑：判断+创建几何集
       ' ======================
       Dim targetHB As HybridBody
       On Error Resume Next ' 捕获“几何集不存在”的错误
       Set targetHB = g_copiedPart.HybridBodies.Item(targetPartName)
       On Error GoTo 0 ' 恢复错误处理
       
       ' 几何集不存在 → 新建
       If targetHB Is Nothing Then
           Set targetHB = g_copiedPart.HybridBodies.Add()
           targetHB.Name = targetPartName
       End If
       
       ' 粘贴到目标几何集
       oSel.Add targetHB
       oSel.PasteSpecial "CATPrtResultWithOutLink"
       CATIA.RefreshDisplay = True
       
       ' 释放对象
       Set targetHB = Nothing


End Sub

' ========== 辅助函数（保留原逻辑） ==========
Private Function IsMeasurableFace(ByVal partDoc As PartDocument, ByVal ref As Reference) As Boolean
    Dim meas As Measurable
    Set meas = partDoc.GetWorkbench("SPAWorkbench").GetMeasurable(ref)
    IsMeasurableFace = (Not meas Is Nothing) And (meas.area > 0)
End Function

Private Function CleanInvalidChars(str As String) As String
    Dim invalidChars As Variant: invalidChars = Array("/", "\", ":", "*", "?", """", "<", ">", "|", " ", "(", ")", "[", "]")
    Dim i As Integer
    For i = 0 To UBound(invalidChars)
        str = Replace(str, invalidChars(i), "_")
    Next
    Do While InStr(str, "__") > 0: str = Replace(str, "__", "_"): Loop
    str = Trim(str)
    If Right(str, 1) = "_" Then str = Left(str, Len(str) - 1)
    If Left(str, 1) = "_" Then str = Mid(str, 2)
    CleanInvalidChars = str
End Function

