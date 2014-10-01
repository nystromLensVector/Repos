Module Module1

    Public SHData As New DataSet1.SHDataTableDataTable
    Public row As DataSet1.SHDataTableRow

    Public HeadData As New DataSet1.Header_DataTableDataTable
    Public hrow As DataSet1.Header_DataTableRow

    Public I10D As New DataSet1.DataTable_XYDataTable
    Public I10row As DataSet1.DataTable_XYRow

    Public FreqPower As Double

    Sub Main()

        Dim i As Integer
        Dim rid As String

        For i = 0 To 15
            row = SHData.NewRow
            row.Index = i
            SHData.Rows.Add(row)
        Next

        For i = 0 To 2
            I10row = I10D.NewRow
            I10row.Xpoint = 0
            I10row.Ypoint = 0
            I10D.Rows.Add(I10row)
        Next

        I10D(0).Xpoint = 0
        I10D(0).Ypoint = 10

        I10D(1).Xpoint = 1000
        I10D(1).Ypoint = 10

        I10D(2).Xpoint = 1000
        I10D(2).Ypoint = 0


        hrow = HeadData.NewRow
        HeadData.Rows.Add(hrow)

        'Dim awrk = From wrk In DB_MJN.ALF_Analysis Where wrk.StartDateTime.Value.DayOfYear > 230 Select rec_id = wrk.Record_ID

        If 1 = 0 Then

            Dim awrk = From wrk In DB_Con.v_Analysis_Lists _
                       Where wrk.DateDif < 21 _
                       Select rec_id = wrk.Record_ID
            Console.WriteLine(String.Format("Record Count = {0}", awrk.Count))

            For Each rec_id In awrk
                rid = rec_id
                'rid = "000011558D"
                If Not (From r In DB_Con.v_ALF_Data_Is_Nulls _
                           Where r.Record_ID.Equals(rid) _
                           Select r).Any() Then
                    'Console.WriteLine(String.Format("record id = {0}", rec_id))
                    'Process_Record(rec_id)
                    Process_Record(rid)
                End If
            Next

        End If

        If 1 = 1 Then

            Dim bwrk = From wrk In DB_Con.v_Analysis_Work_Lists Select wrk
            Console.WriteLine(String.Format("MES Record Count = {0}", bwrk.Count))

            For Each rq In bwrk
                rid = rq.Record_ID
                FreqPower = rq.FreqPower
                If Not (From r In DB_Con.v_ALF_Data_Is_Nulls _
                        Where r.Record_ID.Equals(rid) _
                        Select r).Any() Then
                    Process_Record_MES(rid)
                End If
            Next


        End If

        If AUDIT Then
            Console.WriteLine("End?")
            Console.ReadLine()
        End If


    End Sub

    Sub Process_Record(ByVal RecordNumber As Object)

        Dim rec_str As String
        'Dim rc As Integer = 0

        rec_str = RecordNumber
        'MsgBox([String].Format("Rec = {0}", rec_str), MsgBoxStyle.OkOnly)

        Dim recordSelect = From rec In DB_Con.ALF_Datas _
                Where rec.Record_ID = rec_str _
                Select rec

        'Me.TextBox_RecCount.Text = recordSelect.Count.ToString("D")
        'MsgBox([String].Format("RecordCount = {0}", recordSelect.Count), MsgBoxStyle.OkOnly)

        For Each rec In recordSelect
            'MsgBox([String].Format("RecordID = {0}; PartNumber = {1}", rec.Record_ID, rec.PartNumber), MsgBoxStyle.OkOnly)
            Load_SH_Data(rec)
            Console.WriteLine("Record ID = {0}, Part Number = {1}", rec.Record_ID, rec.PartNumber)
            'rc = rc + 1
            'Me.TextBox_ProcCount.Text = rc.ToString("D")
        Next

    End Sub

    Sub Process_Record_MES(ByVal RecordNumber As Object)

        Dim rec_str As String
        'Dim rc As Integer = 0

        rec_str = RecordNumber
        'MsgBox([String].Format("Rec = {0}", rec_str), MsgBoxStyle.OkOnly)

        Dim recordSelect = From rec In DB_Con.ALF_Data_MEs _
                Where rec.Record_ID = rec_str _
                Select rec

        'Me.TextBox_RecCount.Text = recordSelect.Count.ToString("D")
        'MsgBox([String].Format("RecordCount = {0}", recordSelect.Count), MsgBoxStyle.OkOnly)

        For Each rec In recordSelect
            'MsgBox([String].Format("RecordID = {0}; PartNumber = {1}", rec.Record_ID, rec.PartNumber), MsgBoxStyle.OkOnly)
            Console.WriteLine("Start Record ID = {0}, Part Number = {1}", rec.Record_ID, rec.PartNumber)
            Load_SH_Data_MES(rec)
            Console.WriteLine("Record ID = {0}, Part Number = {1}", rec.Record_ID, rec.PartNumber)
            'rc = rc + 1
            'Me.TextBox_ProcCount.Text = rc.ToString("D")
        Next

    End Sub

    Private Sub Load_SH_Data(ByVal rec As ALF_Data)

        If DB_Update_Head(rec.Record_ID) = 1 Then
            Load_SH_Record(rec, SHData)
            Process_SH_Data(rec, SHData)
            'DB_Update(HeadData)
        Else
            Console.WriteLine("Header Write Error: Record ID = {0}, Part Number = {1}", rec.Record_ID, rec.PartNumber)
        End If



    End Sub

    Private Sub Load_SH_Data_MES(ByVal rec As ALF_Data_ME)

        If DB_Update_Head_MES(rec.Record_ID) = 1 Then
            Load_SH_Record_MES(rec, SHData)
            Process_SH_Data_MES(rec, SHData)
            'DB_Update(HeadData)
        Else
            Console.WriteLine("Header Write Error: Record ID = {0}, Part Number = {1}", rec.Record_ID, rec.PartNumber)
        End If



    End Sub

End Module

