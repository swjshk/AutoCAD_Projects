Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.Windows
Imports Newtonsoft
Imports System.Drawing
Imports System.IO

Public Class EasyFunctionClass
    <CommandMethod("HelloWorld")>
    Public Sub HelloWorld()
        'Test the netload
        'after load the AutoCAD_Projects.dll file, the "HelloWorld" should become a command in AutoCAD
        'and you can type the command and on the Editor page it will show "HelloWorld" msg

        Dim ac_doc As Document = Application.DocumentManager.MdiActiveDocument
        ac_doc.Editor.WriteMessage("HelloWorld!")

    End Sub
    <CommandMethod("PositionSizeAppWindow")>
    Public Sub PositionSizeAppWindow()
        'position the AutoCAD app main windows to a location points
        'size the AutoCAD app main windows by the windows size

        Dim point_app As Windows.Point = New Windows.Point(0, 0)
        Application.MainWindow.DeviceIndependentLocation = point_app

        Dim size_app As Windows.Size = New Windows.Size(400, 400)
        Application.MainWindow.DeviceIndependentSize = size_app

    End Sub
    <CommandMethod("MinimizeWindow")>
    Public Sub MinimizeWindow()
        'minimize the app window
        Core.Application.MainWindow.WindowState = Window.State.Minimized
        MsgBox("Minimized", MsgBoxStyle.SystemModal, "MinMax")
    End Sub
    <CommandMethod("WindowsState")>
    Public Sub WindowsState()
        'show autocad windows state in a msg box
        System.Windows.Forms.MessageBox.Show(Application.MainWindow.WindowState.ToString, "windows state")
    End Sub
    <CommandMethod("SizeDocWindow")>
    Public Sub SizeDocWindow()
        'change CAD drawing window
        Dim active_doc As Document = Application.DocumentManager.MdiActiveDocument
        active_doc.Window.DeviceIndependentLocation = New System.Windows.Point(3, 3)
        active_doc.Window.DeviceIndependentSize = New System.Windows.Size(500, 500)
        'Dim ac_doc_manager As DocumentCollection = Application.DocumentManager
        'Dim ac_doc As Document = DocumentCollectionExtension.Add(ac_doc_manager, "acad.dwt")
        'ac_doc_manager.MdiActiveDocument = ac_doc
    End Sub

End Class

Public Class DocOperation
    <CommandMethod("CreateDrawing")>
    Public Sub CreateDrawing()
        Dim ac_doc_mananger As DocumentCollection = Application.DocumentManager
        Dim ac_doc As Document = DocumentCollectionExtension.Add(ac_doc_mananger, "drawing_a.dwg")
    End Sub
    <CommandMethod("cmd1", CommandFlags.Session)>
    Public Sub cmd1()
        'DocumentCollectionExtension.CloseAll(Application.DocumentManager)

        Dim doc_collection As DocumentCollection = Application.DocumentManager
        For Each doc As Document In doc_collection
            doc.SendStringToExecute("_POINT 1,1,0 ", False, False, True)
            'doc.SendStringToExecute("\x03\x03", False, False, True)
            'DocumentExtension.CloseAndDiscard(doc)
        Next
    End Sub
End Class

Public Class DBOperation
    <CommandMethod("ShowBlkTbl")>
    Public Sub ShowBlkTlb()
        Dim ac_doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim ac_db As Database = ac_doc.Database
        ac_doc.Editor.WriteMessage(ac_doc.Name)
        Using ac_trans As Transaction = ac_db.TransactionManager.StartTransaction()
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'get all the blocktable records from database block table
            Dim btr_list As List(Of BtRecord) = GetBtRecordList(ac_trans, ac_db)
            Dim btrlist_json As String = Json.JsonConvert.SerializeObject(btr_list, Json.Formatting.Indented)
            File.WriteAllText("c:\code_output\btrlist.txt", btrlist_json)

            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ''get all objects from the blocktable record *Model_Space
            'Dim prompt_string_options As PromptStringOptions = New PromptStringOptions(vbLf + "enter the btr name")
            'prompt_string_options.AllowSpaces = True
            'Dim btr_name As String = ac_doc.Editor.GetString(prompt_string_options).StringResult

            'Dim btr_object_list As List(Of BtrObject) = GetBtrObjectsIDLists(ac_trans, ac_db, btr_name)
            'Dim btrobjectlist_json As String = Json.JsonConvert.SerializeObject(btr_object_list, Json.Formatting.Indented)
            'File.WriteAllText("c:\code_output\btrobjectlist.txt", btrobjectlist_json)

            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ''get all the block references objects 
            'Dim blk_ref_list As List(Of BtrObject) = GetBlockReferenceObjectList(btr_object_list)
            'Dim btr_reflist_json As String = Json.JsonConvert.SerializeObject(blk_ref_list, Json.Formatting.Indented)
            'File.WriteAllText("c:\code_output\BlkRefList.txt", btr_reflist_json)

            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'TEST Function GetABlockTableRecord()
            ''get btr name from the prompt string option
            'Dim prompt_string_options As PromptStringOptions = New PromptStringOptions(vbLf + "enter the btr name")
            'prompt_string_options.AllowSpaces = True
            'Dim btr_name As String = ac_doc.Editor.GetString(prompt_string_options).StringResult
            ''get the btr 
            'Dim btr As BlockTableRecord = GetABlockTableRecord(ac_db, ac_trans, btr_name)
            'If btr = Nothing Then
            '    ac_doc.Editor.WriteMessage("There is no " + btr_name)
            'Else
            '    ac_doc.Editor.WriteMessage("The btr: " + btr_name + " exists")
            'End If

            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'TEST Function AddBrOfBtr1ToBtr2()
            'Dim prompt_string_options As PromptStringOptions = New PromptStringOptions(vbLf + "enter the btr name")
            'prompt_string_options.AllowSpaces = True
            'Dim btr1_name As String = ac_doc.Editor.GetString(prompt_string_options).StringResult
            'Dim btr2_name As String = "*Model_Space"
            'Dim insert_point As New Point3d(0, 0, 0)

            'If BtHasBtr(ac_db, ac_trans, btr1_name) And BtHasBtr(ac_db, ac_trans, btr2_name) Then
            '    Dim obj_id As ObjectId = AddBrOfBtr1ToBtr2(ac_db, ac_trans, btr1_name, btr2_name, insert_point)
            'Else
            '    ac_doc.Editor.WriteMessage("There is no " + btr1_name + btr2_name)
            'End If

            Dim br_id_list As List(Of ObjectId) = GetBlockReferenceObjectId(ac_db, ac_trans, "*Model_Space")
            Dim Pawsblocks As List(Of PawsBlockInfo) = GetPawsBlockList(ac_db, ac_trans, br_id_list)
            Dim Pawsblock_json As String = Json.JsonConvert.SerializeObject(Pawsblocks, Json.Formatting.Indented)
            File.WriteAllText("c:\code_output\pawsblockList.txt", Pawsblock_json)
        End Using
    End Sub

    Public Function GetBtRecordList(ByRef ac_trans As Transaction, ByRef ac_db As Database) As List(Of BtRecord)
        Dim btr_list As List(Of BtRecord) = New List(Of BtRecord)
        Dim blk_table As BlockTable = ac_trans.GetObject(ac_db.BlockTableId, OpenMode.ForRead)
        Dim btr As BlockTableRecord = New BlockTableRecord
        For Each btr_id As ObjectId In blk_table
            Dim btrdata As BtRecord = New BtRecord
            btrdata.btr_id = btr_id
            btr = ac_trans.GetObject(btr_id, OpenMode.ForRead)
            btrdata.btr_name = btr.Name
            btrdata.btr_class = btr_id.ObjectClass.Name
            btr_list.Add(btrdata)
        Next
        Return btr_list
    End Function
    Public Function GetBtrObjectsIDLists(ByRef ac_trans As Transaction, ByRef ac_db As Database, ByRef btr_name As String) As List(Of BtrObject)
        Dim object_list As List(Of BtrObject) = New List(Of BtrObject)
        Dim blk_table As BlockTable = ac_trans.GetObject(ac_db.BlockTableId, OpenMode.ForRead)
        Dim btr As BlockTableRecord = ac_trans.GetObject(blk_table(btr_name), OpenMode.ForRead)
        For Each object_id In btr
            Dim btr_object As BtrObject = New BtrObject
            btr_object.object_id = object_id
            btr_object.object_class = object_id.ObjectClass.Name
            If (btr_object.object_class = "AcDbBlockReference") Then
                Dim blk_ref As BlockReference = ac_trans.GetObject(object_id, OpenMode.ForRead)
                Dim ref_btr As BlockTableRecord = ac_trans.GetObject(blk_ref.BlockTableRecord, OpenMode.ForRead)
                btr_object.blk_reference_name = ref_btr.Name
            End If
            object_list.Add(btr_object)

        Next

        Return object_list
    End Function
    Public Function GetBlockReferenceObjectList(ByRef btr_object_list As List(Of BtrObject)) As List(Of BtrObject)
        Dim blk_ref_object_list As New List(Of BtrObject)
        For Each btr_object As BtrObject In btr_object_list
            If (btr_object.object_class = "AcDbBlockReference") Then
                blk_ref_object_list.Add(btr_object)
            End If
        Next
        Return blk_ref_object_list
    End Function
    Public Function GetBlockReferenceObjectId(ByVal db As Database, ByVal tr As Transaction, ByVal btr_name As String) As List(Of ObjectId)
        Dim bt As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
        Dim btr As BlockTableRecord = tr.GetObject(bt.Item(btr_name), OpenMode.ForRead)

        Dim obj_id_list As New List(Of ObjectId)
        For Each obj_id In btr
            If obj_id.ObjectClass.Name = "AcDbBlockReference" Then
                obj_id_list.Add(obj_id)

            End If
        Next

        Return obj_id_list
    End Function
    Public Function GetPawsBlockList(ByVal db As Database, ByVal tr As Transaction, ByVal blk_ref_list As List(Of ObjectId)) As List(Of PawsBlockInfo)
        Dim pawsblock_list As New List(Of PawsBlockInfo)

        For Each obj_id In blk_ref_list
            Dim blk_ref As BlockReference = tr.GetObject(obj_id, OpenMode.ForRead)
            Dim pawsblock As New PawsBlockInfo
            With pawsblock
                .bay_cell = ""
                .device_name = ""
                .pin_number = New List(Of String)
            End With

            For Each attr_ref_id In blk_ref.AttributeCollection
                Dim attr_ref As AttributeReference = tr.GetObject(attr_ref_id, OpenMode.ForRead)
                If attr_ref.Tag = "DEVICE_LOCATION" Then
                    pawsblock.bay_cell = attr_ref.TextString
                End If
                If attr_ref.Tag = "DEVICE_NAME" Then
                    pawsblock.device_name = attr_ref.TextString
                End If
                If attr_ref.Tag Like "TERMINAL_NUMBER*" Then
                    pawsblock.pin_number.Add(attr_ref.TextString)
                End If
            Next
            If pawsblock.bay_cell <> "" And pawsblock.device_name <> "" Then
                pawsblock_list.Add(pawsblock)
            End If
        Next

        Return pawsblock_list
    End Function
    Public Function GetABlockTableRecord(ByVal db As Database, ByVal ac_trans As Transaction, ByVal btr_name As String) As BlockTableRecord
        Dim bt As BlockTable = ac_trans.GetObject(db.BlockTableId, OpenMode.ForRead)
        If bt.Has(btr_name) Then
            Dim btr As BlockTableRecord = ac_trans.GetObject(bt(btr_name), OpenMode.ForRead)
            Return btr
        Else
            Return Nothing
        End If
    End Function
    Public Function GetModelSpace(ByVal db As Database, ByVal ac_trans As Transaction) As BlockTableRecord
        Dim bt As BlockTable = ac_trans.GetObject(db.BlockTableId, OpenMode.ForRead)
        Dim btr As BlockTableRecord = ac_trans.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForRead)
        Return btr
    End Function
    Public Function BtHasBtr(ByRef db As Database, ByRef ac_trans As Transaction, ByVal btr_name As String) As Boolean
        Dim bt As BlockTable = ac_trans.GetObject(db.BlockTableId, OpenMode.ForRead)
        Dim btr_exist As Boolean = 0
        If bt.Has(btr_name) Then
            btr_exist = 1
        End If
        Return btr_exist
    End Function
    Public Function AddBrOfBtr1ToBtr2(ByRef db As Database, ByRef ac_trans As Transaction, ByVal btr1_name As String, ByVal btr2_name As String, ByVal insert_point As Point3d) As ObjectId
        'add a BlockReference of BlockTableRecord1 to BlockTableRecord2 at a certain point
        'check both btr1 and btr2 exist via BthasBtr() before using this function

        Dim bt As BlockTable = ac_trans.GetObject(db.BlockTableId, OpenMode.ForRead)
        Dim btr1 As BlockTableRecord = ac_trans.GetObject(bt.Item(btr1_name), OpenMode.ForRead)
        Dim br1 As New BlockReference(insert_point, bt.Item(btr1_name))
        Dim btr2 As BlockTableRecord = ac_trans.GetObject(bt.Item(btr2_name), OpenMode.ForWrite)
        Dim obj_id As ObjectId = btr2.AppendEntity(br1) '!!!insert the br to the btr first then attach the attributes!!!!!
        ac_trans.AddNewlyCreatedDBObject(br1, True) '!Attention! Need to confirm the addition

        If btr1.HasAttributeDefinitions Then
            For Each btr1_obj_id As ObjectId In btr1
                If btr1_obj_id.ObjectClass.Name = "AcDbAttributeDefinition" Then
                    Dim attr_def As AttributeDefinition = ac_trans.GetObject(btr1_obj_id, OpenMode.ForRead)
                    Using attr_ref As New AttributeReference
                        attr_ref.SetAttributeFromBlock(attr_def, br1.BlockTransform)
                        attr_ref.Position = attr_def.Position.TransformBy(br1.BlockTransform)
                        attr_ref.TextString = attr_def.TextString
                        br1.AttributeCollection.AppendAttribute(attr_ref)
                        ac_trans.AddNewlyCreatedDBObject(attr_ref, True)
                    End Using
                End If
            Next
        End If

        ac_trans.Commit()

        Return obj_id
    End Function

End Class

Public Class BtRecord
    Public btr_id As ObjectId
    Public btr_name As String
    Public btr_class As String

End Class
Public Class BtrObject
    Public object_id As ObjectId
    Public object_class As String
    Public blk_reference_name As String
End Class
Public Class AcBlockRefInfo
    Public object_id As ObjectId
    Public block_name As String
    Public attr_collection As AttributeCollection
End Class

Public Class PawsBlockInfo
    'Public device_id As String
    Public device_name As String ' eg."BPTB"
    Public bay_cell As String 'eg."[1D3]"
    Public pin_number As List(Of String)
    Public Object_id As String 'ObjectId.Handle eg."51AC"
    Public doc_name As String 'drawing file name eg."E37019765-056-04-3.DWG"
    Public doc_directory As String 'drawing file directory eg. "C:\FO\37839482\"
    Public SQD_PN As String 'eg. "9080GR6"
End Class

Public Class DeviceInfo
    Public device_id As String
    Public device_name As String
    Public bay_cell As String
    Public pins_number As List(Of String)
    Public SQD_PN As String
    Public MANUFACTURER As String
    Public MODEL As String
End Class

Public Class KeyForPawsBlock
    Public device_name As String
    Public bay_cell As String
End Class
