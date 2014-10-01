
Imports System.Math

Public Module MyConstants
    Public Const VERSION As Integer = 3
    Public Const ZERO_VOLT As Double = 3  ' No longer used see Zero_Volt1 below MJN 20120608
    Public Const ZERO_FREQ As Double = 0.12
    Public Const INF_FREQ As Double = 1

    Public Const AUDIT As Boolean = False
End Module

Public Module DB_Analysis


    Public DB_Con As New Con_DB_DataContext()

    'Public DB_MJN As New MJN_DB_DataContext()

    Public Data_Int As New DataSet2.ResultsTableDataTable
    Public drw As DataSet2.ResultsTableRow

    Public Zero_Volt1 As Double = 3.0  ' Added to replace ZERO_VOLT constant MJN 20120608

    Public Sub Load_SH_Record(ByVal rec As ALF_Data, ByRef SHData As DataSet1.SHDataTableDataTable)

        Dim row As DataSet1.SHDataTableRow

        For Each row In SHData.Rows
            Select Case row.Index
                Case 0
                    row.Volt = rec._42
                    row.Freq = rec._58
                    row.Power_H = rec._74
                    row.Power_V = rec._90
                    row.RMS_H = rec._106
                    row.RMS_V = rec._122
                    row.Power_avg = rec._138
                    row.RMS_avg = rec._154
                Case 1
                    row.Volt = rec._43
                    row.Freq = rec._59
                    row.Power_H = rec._75
                    row.Power_V = rec._91
                    row.RMS_H = rec._107
                    row.RMS_V = rec._123
                    row.Power_avg = rec._139
                    row.RMS_avg = rec._155
                Case 2
                    row.Volt = rec._44
                    row.Freq = rec._60
                    row.Power_H = rec._76
                    row.Power_V = rec._92
                    row.RMS_H = rec._108
                    row.RMS_V = rec._124
                    row.Power_avg = rec._140
                    row.RMS_avg = rec._156
                Case 3
                    row.Volt = rec._45
                    row.Freq = rec._61
                    row.Power_H = rec._77
                    row.Power_V = rec._93
                    row.RMS_H = rec._109
                    row.RMS_V = rec._125
                    row.Power_avg = rec._141
                    row.RMS_avg = rec._157
                Case 4
                    row.Volt = rec._46
                    row.Freq = rec._62
                    row.Power_H = rec._78
                    row.Power_V = rec._94
                    row.RMS_H = rec._110
                    row.RMS_V = rec._126
                    row.Power_avg = rec._142
                    row.RMS_avg = rec._158
                Case 5
                    row.Volt = rec._47
                    row.Freq = rec._63
                    row.Power_H = rec._79
                    row.Power_V = rec._95
                    row.RMS_H = rec._111
                    row.RMS_V = rec._127
                    row.Power_avg = rec._143
                    row.RMS_avg = rec._159
                Case 6
                    row.Volt = rec._48
                    row.Freq = rec._64
                    row.Power_H = rec._80
                    row.Power_V = rec._96
                    row.RMS_H = rec._112
                    row.RMS_V = rec._128
                    row.Power_avg = rec._144
                    row.RMS_avg = rec._160
                Case 7
                    row.Volt = rec._49
                    row.Freq = rec._65
                    row.Power_H = rec._81
                    row.Power_V = rec._97
                    row.RMS_H = rec._113
                    row.RMS_V = rec._129
                    row.Power_avg = rec._145
                    row.RMS_avg = rec._161
                Case 8
                    row.Volt = rec._50
                    row.Freq = rec._66
                    row.Power_H = rec._82
                    row.Power_V = rec._98
                    row.RMS_H = rec._114
                    row.RMS_V = rec._130
                    row.Power_avg = rec._146
                    row.RMS_avg = rec._162
                Case 9
                    row.Volt = rec._51
                    row.Freq = rec._67
                    row.Power_H = rec._83
                    row.Power_V = rec._99
                    row.RMS_H = rec._115
                    row.RMS_V = rec._131
                    row.Power_avg = rec._147
                    row.RMS_avg = rec._163
                Case 10
                    row.Volt = rec._52
                    row.Freq = rec._68
                    row.Power_H = rec._84
                    row.Power_V = rec._100
                    row.RMS_H = rec._116
                    row.RMS_V = rec._132
                    row.Power_avg = rec._148
                    row.RMS_avg = rec._164
                Case 11
                    row.Volt = rec._53
                    row.Freq = rec._69
                    row.Power_H = rec._85
                    row.Power_V = rec._101
                    row.RMS_H = rec._117
                    row.RMS_V = rec._133
                    row.Power_avg = rec._149
                    row.RMS_avg = rec._165
                Case 12
                    row.Volt = rec._54
                    row.Freq = rec._70
                    row.Power_H = rec._86
                    row.Power_V = rec._102
                    row.RMS_H = rec._118
                    row.RMS_V = rec._134
                    row.Power_avg = rec._150
                    row.RMS_avg = rec._166
                Case 13
                    row.Volt = rec._55
                    row.Freq = rec._71
                    row.Power_H = rec._87
                    row.Power_V = rec._103
                    row.RMS_H = rec._119
                    row.RMS_V = rec._135
                    row.Power_avg = rec._151
                    row.RMS_avg = rec._167
                Case 14
                    row.Volt = rec._56
                    row.Freq = rec._72
                    row.Power_H = rec._88
                    row.Power_V = rec._104
                    row.RMS_H = rec._120
                    row.RMS_V = rec._136
                    row.Power_avg = rec._152
                    row.RMS_avg = rec._168
                Case 15
                    row.Volt = rec._57
                    row.Freq = rec._73
                    row.Power_H = rec._89
                    row.Power_V = rec._105
                    row.RMS_H = rec._121
                    row.RMS_V = rec._137
                    row.Power_avg = rec._153
                    row.RMS_avg = rec._169
            End Select
        Next

    End Sub

    Public Sub Load_SH_Record_MES(ByVal rec As ALF_Data_ME, ByRef SHData As DataSet1.SHDataTableDataTable)

        Dim row As DataSet1.SHDataTableRow

        For Each row In SHData.Rows
            Select Case row.Index
                Case 0
                    row.Volt = rec._42
                    row.Freq = rec._58
                    row.Power_H = rec._74
                    row.Power_V = rec._90
                    row.RMS_H = rec._106
                    row.RMS_V = rec._122
                    row.Power_avg = rec._138
                    row.RMS_avg = rec._154
                Case 1
                    row.Volt = rec._43
                    row.Freq = rec._59
                    row.Power_H = rec._75
                    row.Power_V = rec._91
                    row.RMS_H = rec._107
                    row.RMS_V = rec._123
                    row.Power_avg = rec._139
                    row.RMS_avg = rec._155
                Case 2
                    row.Volt = rec._44
                    row.Freq = rec._60
                    row.Power_H = rec._76
                    row.Power_V = rec._92
                    row.RMS_H = rec._108
                    row.RMS_V = rec._124
                    row.Power_avg = rec._140
                    row.RMS_avg = rec._156
                Case 3
                    row.Volt = rec._45
                    row.Freq = rec._61
                    row.Power_H = rec._77
                    row.Power_V = rec._93
                    row.RMS_H = rec._109
                    row.RMS_V = rec._125
                    row.Power_avg = rec._141
                    row.RMS_avg = rec._157
                Case 4
                    row.Volt = rec._46
                    row.Freq = rec._62
                    row.Power_H = rec._78
                    row.Power_V = rec._94
                    row.RMS_H = rec._110
                    row.RMS_V = rec._126
                    row.Power_avg = rec._142
                    row.RMS_avg = rec._158
                Case 5
                    row.Volt = rec._47
                    row.Freq = rec._63
                    row.Power_H = rec._79
                    row.Power_V = rec._95
                    row.RMS_H = rec._111
                    row.RMS_V = rec._127
                    row.Power_avg = rec._143
                    row.RMS_avg = rec._159
                Case 6
                    row.Volt = rec._48
                    row.Freq = rec._64
                    row.Power_H = rec._80
                    row.Power_V = rec._96
                    row.RMS_H = rec._112
                    row.RMS_V = rec._128
                    row.Power_avg = rec._144
                    row.RMS_avg = rec._160
                Case 7
                    row.Volt = rec._49
                    row.Freq = rec._65
                    row.Power_H = rec._81
                    row.Power_V = rec._97
                    row.RMS_H = rec._113
                    row.RMS_V = rec._129
                    row.Power_avg = rec._145
                    row.RMS_avg = rec._161
                Case 8
                    row.Volt = rec._50
                    row.Freq = rec._66
                    row.Power_H = rec._82
                    row.Power_V = rec._98
                    row.RMS_H = rec._114
                    row.RMS_V = rec._130
                    row.Power_avg = rec._146
                    row.RMS_avg = rec._162
                Case 9
                    row.Volt = rec._51
                    row.Freq = rec._67
                    row.Power_H = rec._83
                    row.Power_V = rec._99
                    row.RMS_H = rec._115
                    row.RMS_V = rec._131
                    row.Power_avg = rec._147
                    row.RMS_avg = rec._163
                Case 10
                    row.Volt = rec._52
                    row.Freq = rec._68
                    row.Power_H = rec._84
                    row.Power_V = rec._100
                    row.RMS_H = rec._116
                    row.RMS_V = rec._132
                    row.Power_avg = rec._148
                    row.RMS_avg = rec._164
                Case 11
                    row.Volt = rec._53
                    row.Freq = rec._69
                    row.Power_H = rec._85
                    row.Power_V = rec._101
                    row.RMS_H = rec._117
                    row.RMS_V = rec._133
                    row.Power_avg = rec._149
                    row.RMS_avg = rec._165
                Case 12
                    row.Volt = rec._54
                    row.Freq = rec._70
                    row.Power_H = rec._86
                    row.Power_V = rec._102
                    row.RMS_H = rec._118
                    row.RMS_V = rec._134
                    row.Power_avg = rec._150
                    row.RMS_avg = rec._166
                Case 13
                    row.Volt = rec._55
                    row.Freq = rec._71
                    row.Power_H = rec._87
                    row.Power_V = rec._103
                    row.RMS_H = rec._119
                    row.RMS_V = rec._135
                    row.Power_avg = rec._151
                    row.RMS_avg = rec._167
                Case 14
                    row.Volt = rec._56
                    row.Freq = rec._72
                    row.Power_H = rec._88
                    row.Power_V = rec._104
                    row.RMS_H = rec._120
                    row.RMS_V = rec._136
                    row.Power_avg = rec._152
                    row.RMS_avg = rec._168
                Case 15
                    row.Volt = rec._57
                    row.Freq = rec._73
                    row.Power_H = rec._89
                    row.Power_V = rec._105
                    row.RMS_H = rec._121
                    row.RMS_V = rec._137
                    row.Power_avg = rec._153
                    row.RMS_avg = rec._169
            End Select
        Next

    End Sub

    Public Sub Process_SH_Data(ByVal rec As ALF_Data, ByVal SHData As DataSet1.SHDataTableDataTable)

        Dim HF As Boolean

        drw = Data_Int.NewRow

        drw.Record_ID = rec.Record_ID
        drw.Measurement_Grade = -1

       
        Select Case Strings.Left(rec.PartNumber, 3)
            Case "MHF"
                HF = True
            Case "MFG"
                HF = True
            Case "MFA"
                HF = False
            Case "MFZ"
                HF = False
            Case "MHL"
                HF = True
            Case Else
                HF = False
        End Select

        If HF Then

            Dim FindHPolarquery = From fhpol In SHData _
                                 Where fhpol.Volt > Zero_Volt1 _
                                 Order By fhpol.Power_H Descending _
                                 Select fhpol

            Dim FindVPolarquery = From fvpol In SHData _
                                  Where fvpol.Volt > Zero_Volt1 _
                                  Order By fvpol.Power_V Descending _
                                  Select fvpol

            'MsgBox([String].Format("Ph = {0}; Fh = {1}; Pv = {2}; Fv = {3}", FindHPolarquery.First.Power_H, FindHPolarquery.First.Freq, FindVPolarquery.First.Power_V, FindVPolarquery.First.Freq), MsgBoxStyle.OkOnly)

            If FindHPolarquery.First.Power_H > FindVPolarquery.First.Power_V Then
                Analyze_Polar(drw, SHData, 1, "H")
                Analyze_Polar(drw, SHData, 0, "H")
            Else
                Analyze_Polar(drw, SHData, 2, "V")
                Analyze_Polar(drw, SHData, 0, "V")
            End If
        Else
            Analyze_Polar(drw, SHData, 1, "H")
            Analyze_Polar(drw, SHData, 2, "V")
            Analyze_Polar(drw, SHData, 0, "A")

            If DB_Update(drw, Analyze_Pol_Offset(SHData), "Pol_Off", 0) < 1 Then PrintError(drw.Record_ID, "Pol_Off")
            If DB_Update(drw, Analyze_Max_Pol_Offset(SHData), "Max_Pol_Off", 0) < 1 Then PrintError(drw.Record_ID, "Max_Pol_Off")
            If DB_Update(drw, Analyze_Max_Pol_Offset2(SHData), "Max_Pol_Off2", 0) < 1 Then PrintError(drw.Record_ID, "Max_Pol_Off2")
            If Val_Measure_Pol(SHData, "H") And Val_Measure_Pol(SHData, "V") And Val_Measure_Pol(SHData, "A") Then
                drw.Measurement_Grade = 1
            End If
        End If
        If DB_Update2(drw) < 1 Then PrintError(drw.Record_ID, "Data write error")
    End Sub

    Public Sub Process_SH_Data_MES(ByVal rec As ALF_Data_ME, ByVal SHData As DataSet1.SHDataTableDataTable)

        Dim HF As Boolean

        drw = Data_Int.NewRow

        drw.Record_ID = rec.Record_ID
        drw.Measurement_Grade = -1

        ' Added code to change the threasehold voltage for determining the zero volt measurement based on the max voltage
        ' MJN 20120608

        Dim MaxVoltQuery = (From v In SHData _
                   Order By v.Volt Descending _
                   Select v.Volt).Take(1)


        If MaxVoltQuery.First > 20 Then Zero_Volt1 = 12 Else Zero_Volt1 = 3

        'MsgBox([String].Format("Rec = {0}, Max Volt = {1}, Volt Div = {2}", rec.Record_ID, MaxVoltQuery.First, Zero_Volt1), MsgBoxStyle.OkOnly)

        ' End new code MJN 20120608

        Select Case Strings.Left(rec.PartNumber, 2)
            Case "FL"
                HF = False
            Case Else
                HF = False
        End Select

        If HF Then

            Dim FindHPolarquery = From fhpol In SHData _
                                 Where fhpol.Volt > Zero_Volt1 _
                                 Order By fhpol.Power_H Descending _
                                 Select fhpol

            Dim FindVPolarquery = From fvpol In SHData _
                                  Where fvpol.Volt > Zero_Volt1 _
                                  Order By fvpol.Power_V Descending _
                                  Select fvpol

            'MsgBox([String].Format("Ph = {0}; Fh = {1}; Pv = {2}; Fv = {3}", FindHPolarquery.First.Power_H, FindHPolarquery.First.Freq, FindVPolarquery.First.Power_V, FindVPolarquery.First.Freq), MsgBoxStyle.OkOnly)

            If FindHPolarquery.First.Power_H > FindVPolarquery.First.Power_V Then
                Analyze_Polar(drw, SHData, 1, "H")
                Analyze_Polar(drw, SHData, 0, "H")
            Else
                Analyze_Polar(drw, SHData, 2, "V")
                Analyze_Polar(drw, SHData, 0, "V")
            End If
        Else
            Analyze_Polar(drw, SHData, 1, "H")
            Analyze_Polar(drw, SHData, 2, "V")
            Analyze_Polar(drw, SHData, 0, "A")

            If DB_Update(drw, Analyze_Pol_Offset(SHData), "Pol_Off", 0) < 1 Then PrintError(drw.Record_ID, "Pol_Off")
            If DB_Update(drw, Analyze_Max_Pol_Offset(SHData), "Max_Pol_Off", 0) < 1 Then PrintError(drw.Record_ID, "Max_Pol_Off")
            If DB_Update(drw, Analyze_Max_Pol_Offset2(SHData), "Max_Pol_Off2", 0) < 1 Then PrintError(drw.Record_ID, "Max_Pol_Off2")
            If Val_Measure_Pol(SHData, "H") And Val_Measure_Pol(SHData, "V") And Val_Measure_Pol(SHData, "A") Then
                drw.Measurement_Grade = 1
            End If
            If DB_Update(drw, Delta_Power(drw.I10Dptinf_Freq_avg, SHData, "H"), "Future_1", 0) < 1 Then PrintError(drw.Record_ID, "Future_1")
            If DB_Update(drw, Delta_Power(drw.I10Dptinf_Freq_avg, SHData, "V"), "Future_2", 0) < 1 Then PrintError(drw.Record_ID, "Future_2")
            If DB_Update(drw, (drw.Future_1 - drw.Future_2), "Future_3", 0) < 1 Then PrintError(drw.Record_ID, "Future_3")
        End If
        If DB_Update2(drw) < 1 Then PrintError(drw.Record_ID, "Data write error")
    End Sub

    Private Function Delta_Power(ByVal Freq_10D As Double, _
                                 ByVal SHData As DataSet1.SHDataTableDataTable, _
                                 ByVal pol As String) As Double

        Dim temp As Double

        Dim Pow_low = (From rw In SHData _
                        Where rw.Volt > Zero_Volt1 _
                        And rw.Freq > ZERO_FREQ _
                        And rw.Freq < Freq_10D _
                        Order By rw.Freq Descending _
                        Select rw).Take(1)

        Dim Pow_high = (From rw In SHData _
                        Where rw.Volt > Zero_Volt1 _
                        And rw.Freq > ZERO_FREQ _
                        And rw.Freq > Freq_10D _
                        Order By rw.Freq Ascending _
                        Select rw).Take(1)

        Dim ROPinfquery = (From rw In SHData _
                          Where rw.Volt > Zero_Volt1 _
                          And rw.Freq > ZERO_FREQ _
                          Order By rw.Freq Ascending _
                          Select rw).Take(1)


        If ((Pow_low.Count > 0) And (Pow_high.Count > 0) And (ROPinfquery.Count > 0)) Then
            Select Case pol
                Case "H"


                    temp = MJNFitting.Fitting.Interpolate( _
                                    (Pow_low.First.Power_H - ROPinfquery.First.Power_H), _
                                      (Pow_high.First.Power_H - ROPinfquery.First.Power_H), _
                                      Pow_low.First.Freq, _
                                      Pow_high.First.Freq, Freq_10D)
                    If AUDIT Then Console.WriteLine("H: P0 = {0}, P1 = {1}, F0 = {2}, F1 = {3}, Fx = {4}, Px = {5}", _
                                      (Pow_low.First.Power_H - ROPinfquery.First.Power_H), _
                                      (Pow_high.First.Power_H - ROPinfquery.First.Power_H), _
                                      Pow_low.First.Freq, _
                                      Pow_high.First.Freq, Freq_10D, temp)
                Case "V"

                    temp = MJNFitting.Fitting.Interpolate( _
                                    (Pow_low.First.Power_V - ROPinfquery.First.Power_V), _
                                      (Pow_high.First.Power_V - ROPinfquery.First.Power_V), _
                                      Pow_low.First.Freq, _
                                      Pow_high.First.Freq, Freq_10D)
                    If AUDIT Then Console.WriteLine("V: P0 = {0}, P1 = {1}, F0 = {2}, F1 = {3}, Fx = {4}, Px = {5}", _
                                      (Pow_low.First.Power_V - ROPinfquery.First.Power_V), _
                                      (Pow_high.First.Power_V - ROPinfquery.First.Power_V), _
                                      Pow_low.First.Freq, _
                                      Pow_high.First.Freq, Freq_10D, temp)
                Case Else
                    temp = -1
            End Select
        Else
            temp = -1
        End If


        Return temp

    End Function


    Private Sub Analyze_Polar(ByRef drw As DataSet2.ResultsTableRow, ByVal SHData As DataSet1.SHDataTableDataTable, ByVal dr As Integer, ByVal pol As String)

        Dim table As New DataSet1.Pol_DataTableDataTable
        Dim rw As DataSet1.Pol_DataTableRow

        Dim tablequery = From ro In SHData Select ro

        For Each ro In tablequery
            rw = table.NewRow
            With rw
                .Index = ro.Index
                .Volt = ro.Volt
                .Freq = ro.Freq
                Select Case pol
                    Case "H"
                        .power = ro.Power_H
                        .rms = ro.RMS_H
                    Case "V"
                        .power = ro.Power_V
                        .rms = ro.RMS_V
                    Case Else
                        .power = ro.Power_avg
                        .rms = ro.RMS_avg
                End Select
            End With
            table.Rows.Add(rw)
        Next

        'Console.WriteLine(String.Format("table.count = {0}", table.Count))
        Analyze_Pol(drw, dr, table)

    End Sub



    Private Function Analyze_Pol_Offset(ByRef SHData As DataSet1.SHDataTableDataTable)

        Dim peakfreq = 0
        Dim PolOff As Double

        Dim NZtable = From rw In SHData _
                      Where rw.Volt > Zero_Volt1 _
                      And rw.Freq > ZERO_FREQ _
                      Select rw

        Dim Peakquery = From pk In NZtable _
                         Order By pk.Power_avg Descending _
                         Select pk
        peakfreq = Peakquery.First.Freq()

        Dim table = From ro In NZtable _
                    Where ro.Freq < peakfreq _
                    And ro.Freq > INF_FREQ _
                    And ro.Power_H < 12 _
                    And ro.Power_V < 12 _
                    Order By ro.Freq Descending _
                    Select ro.Power_H, ro.Power_V

        PolOff = 0
        For Each ro In table
            PolOff = PolOff + (ro.Power_H - ro.Power_V) ^ 2
        Next

        PolOff = Sqrt(PolOff / table.Count)

        Return PolOff

    End Function

    Private Function Analyze_Max_Pol_Offset(ByRef SHData As DataSet1.SHDataTableDataTable)

        Dim peakfreq = 0
        Dim MaxPolOff As Double
        Dim d_temp As Double

        Dim NZtable = From rw In SHData _
                      Where rw.Volt > Zero_Volt1 _
                      And rw.Freq > ZERO_FREQ _
                      Select rw

        Dim Peakquery = From pk In NZtable _
                         Order By pk.Power_avg Descending _
                         Select pk
        peakfreq = Peakquery.First.Freq()

        Dim table = From ro In NZtable _
                    Where ro.Freq < peakfreq _
                    And ro.Freq > INF_FREQ _
                    And ro.Power_H < 12 _
                    And ro.Power_V < 12 _
                    Order By ro.Freq Descending _
                    Select ro.Power_H, ro.Power_V

        MaxPolOff = 0
        For Each ro In table
            d_temp = Abs(ro.Power_H - ro.Power_V)
            If d_temp > MaxPolOff Then
                MaxPolOff = d_temp
            End If
        Next

        Return MaxPolOff

    End Function

    Private Function Analyze_Max_Pol_Offset2(ByRef SHData As DataSet1.SHDataTableDataTable)

        Dim peakfreq = 0
        Dim MaxPolOff2 As Double
        Dim d_temp As Double

        Dim NZtable = From rw In SHData _
                      Where rw.Volt > Zero_Volt1 _
                      And rw.Freq > ZERO_FREQ _
                      Select rw

        Dim Peakquery = From pk In NZtable _
                         Order By pk.Power_avg Descending _
                         Select pk
        peakfreq = Peakquery.First.Freq()

        Dim table = From ro In NZtable _
                    Where ro.Freq < peakfreq _
                    And ro.Freq > INF_FREQ _
                    And ro.Power_H < 12 _
                    And ro.Power_V < 12 _
                    Order By ro.Freq Descending _
                    Select ro.Power_H, ro.Power_V

        MaxPolOff2 = 0
        For Each ro In table
            d_temp = ro.Power_H - ro.Power_V
            If Abs(d_temp) > Abs(MaxPolOff2) Then
                MaxPolOff2 = d_temp
            End If
        Next

        Return MaxPolOff2

    End Function

    Private Function Diopt_Power(ByVal drw As DataSet2.ResultsTableRow, _
                                ByVal ref As Double, _
                                ByVal key As String, _
                                ByVal dr As Integer, _
                                ByVal nztable As System.Data.EnumerableRowCollection(Of New_Analyze_SH.DataSet1.Pol_DataTableRow)) As Integer

        Dim retval As Integer = 0
        Dim temp As Double

        Dim FreqMinquery = From fmin In nztable _
                           Order By fmin.Freq Ascending _
                           Select fmin

        Dim q1 = (From ro In nztable _
                    Where ro.power >= ref _
                    Order By ro.Freq Ascending _
                    Select ro).Take(1)
        If q1.Count > 0 Then

            If q1.First.Freq > FreqMinquery.First.Freq Then

                Dim q2 = (From int In nztable _
                                Where int.Freq <= q1.First.Freq _
                                Order By int.Freq Descending _
                                Select int).Take(2)

                If q2.Count > 1 Then



                    If DB_Update(drw, q2.First.Volt, key + "_Volt", dr) < 1 Then PrintError(drw.Record_ID, "_Volt")

                    temp = MJNFitting.Fitting.Interpolate(q2(1).Freq, _
                                       q2(0).Freq, _
                                       q2(1).power, _
                                       q2(0).power, ref)

                    If DB_Update(drw, temp, key + "_Freq", dr) < 1 Then PrintError(drw.Record_ID, key + "_Freq")

                    temp = MJNFitting.Fitting.Interpolate(q2(1).rms, _
                                       q2(0).rms, _
                                       q2(1).power, _
                                       q2(0).power, ref)

                    If DB_Update(drw, temp, key + "_RMS", dr) < 1 Then PrintError(drw.Record_ID, key + "_RMS")

                Else

                    retval = -1

                End If

            Else
                retval = -1
            End If
        Else
            retval = -1
        End If

        Return retval

    End Function

    Private Function Analyze_ClearOP(ByVal drw As DataSet2.ResultsTableRow, _
                                    ByVal key As String, _
                                    ByVal dr As Integer, _
                                    ByVal nztable As System.Data.EnumerableRowCollection(Of New_Analyze_SH.DataSet1.Pol_DataTableRow)) As Integer

        Dim retval As Integer = 0
        Dim rmslimit As Double

        Select Case key
            Case "ClearOP25"
                rmslimit = 0.25
            Case "ClearOP20"
                rmslimit = 0.2
            Case "ClearOP15"
                rmslimit = 0.15
        End Select

        Dim Peakquery = From pk In nztable _
                 Order By pk.power Descending _
                 Select pk

        Dim FCq = (From clr In nztable _
          Where clr.Freq <= Peakquery.First.Freq _
          And clr.rms <= rmslimit _
          Order By clr.Freq Descending _
          Select clr).Take(1)

        'MsgBox([String].Format("Freq0 = {0}; Power0 = {1}; RMS0 = {2}", PeakHquery.First.Freq, PeakHquery.First.Power_H, PeakHquery.First.RMS_H), MsgBoxStyle.OkOnly)

        'MsgBox([String].Format("Clear Count = {0}", FindClearHquery.Count), MsgBoxStyle.OkOnly)

        If FCq.Count > 0 Then

            'MsgBox([String].Format("Freq0 = {0}; FreqCl = {1}; rmsCl = {2}", PeakHquery.First.Freq, FindClearHquery.First.Freq, FindClearHquery.First.RMS_H), MsgBoxStyle.OkOnly)
            If DB_Update(drw, FCq.First.power, key + "_Power", dr) < 1 Then PrintError(drw.Record_ID, key + "_Power")
            If DB_Update(drw, FCq.First.rms, key + "_RMS", dr) < 1 Then PrintError(drw.Record_ID, key + "_RMS")
            If DB_Update(drw, FCq.First.Volt, key + "_Volt", dr) < 1 Then PrintError(drw.Record_ID, key + "_Volt")
            If DB_Update(drw, FCq.First.Freq, key + "_Freq", dr) < 1 Then PrintError(drw.Record_ID, key + "_Freq")

        End If

        Return retval

    End Function


    Private Function Val_Measure_Pol(ByVal SHData As DataSet1.SHDataTableDataTable, ByVal pol As String) As Boolean

        Dim table As New DataSet1.Pol_DataTableDataTable
        Dim rw As DataSet1.Pol_DataTableRow

        Dim retval As Boolean = False

        Dim tablequery = From ro In SHData _
                         Where ro.Volt > Zero_Volt1 _
                         And ro.Freq > ZERO_FREQ _
                         Select ro

        For Each ro In tablequery
            rw = table.NewRow
            With rw
                .Index = ro.Index
                .Volt = ro.Volt
                .Freq = ro.Freq
                Select Case pol
                    Case "H"
                        .power = ro.Power_H
                        .rms = ro.RMS_H
                    Case "V"
                        .power = ro.Power_V
                        .rms = ro.RMS_V
                    Case Else
                        .power = ro.Power_avg
                        .rms = ro.RMS_avg
                End Select
            End With
            table.Rows.Add(rw)
        Next

        'Console.WriteLine(String.Format("table.count = {0}", table.Count))
        Return Valid_Measurement_Pol(table)

    End Function

    Private Function Valid_Measurement_Pol(ByVal table As DataSet1.Pol_DataTableDataTable) As Boolean

        Dim retval As Boolean = False
        Dim lastP As Double
        Dim highT, lowT As Boolean



        Dim Peakquery = From pk In table _
                 Order By pk.power Descending _
                 Select pk

        Dim HiFreq = From hf In table _
                     Where hf.Freq > Peakquery.First.Freq _
                     Order By hf.Freq Ascending _
                     Select hf

        Dim LoFreq = From lf In table _
                     Where lf.Freq < Peakquery.First.Freq _
                     Order By lf.Freq Descending _
                     Select lf

        If HiFreq.Count = 2 Or HiFreq.Count = 3 Then
            lastP = Peakquery.First.power
            highT = True
            For Each hf In HiFreq
                If lastP < hf.power Then highT = False
                lastP = hf.power
            Next
            If highT Then
                lastP = Peakquery.First.power
                lowT = True
                For Each lf In LoFreq
                    If lastP < lf.power Then lowT = False
                    lastP = lf.power
                Next
                If lowT Then retval = True
            End If
        End If

        'MsgBox([String].Format("Freq0 = {0}; Power0 = {1}; RMS0 = {2}", PeakHquery.First.Freq, PeakHquery.First.Power_H, PeakHquery.First.RMS_H), MsgBoxStyle.OkOnly)

        'MsgBox([String].Format("Clear Count = {0}", FindClearHquery.Count), MsgBoxStyle.OkOnly)

        Return retval

    End Function

    Private Sub Analyze_Pol(ByRef drw As DataSet2.ResultsTableRow, _
                           ByVal dr As Integer, _
                           ByVal table As DataSet1.Pol_DataTableDataTable)

        Dim ref As Double

        Dim xy(1, 2) As Double
        Dim xr(1, 2) As Double
        Dim coeff_y(2) As Double
        Dim coeff_r(2) As Double

        Dim temp As Double

        'Conditions for ROP0 state

        'ZeroVolt = 12
        'ZeroFreq = 0.12

        'PeakPower = 0
        'PeakFreq = 0

        Dim FreqMaxquery = From fmax In table _
                           Where fmax.Volt > Zero_Volt1 _
                           Order By fmax.Freq Descending _
                           Select fmax

        Dim FreqMinquery = From fmin In table _
                           Where fmin.Volt > Zero_Volt1 _
                           Order By fmin.Freq Ascending _
                           Select fmin

        'MsgBox([String].Format("FreqMax = {0}; FreqMin = {1}", FreqMaxquery.First.Freq, FreqMinquery.First.Freq), MsgBoxStyle.OkOnly)

        'If you have a problem below where the "zero" voltage is higher than ZERO_VOLT
        ' change the "and" to "or" in the call below

        Dim ROP0query = From rop In table _
                        Where rop.Volt < Zero_Volt1 _
                        And rop.Freq < ZERO_FREQ _
                        Order By rop.Index Ascending _
                        Select rop

        If ROP0query.Count > 0 Then
            If DB_Update(drw, ROP0query.First.power, "ROP_0", dr) < 1 Then PrintError(drw.Record_ID, "ROP_0")
            If DB_Update(drw, ROP0query.First.rms, "RMS_0", dr) < 1 Then PrintError(drw.Record_ID, "RMS_0")
        End If
        'MsgBox([String].Format("ROP 0 freq = {0}", ROP0query.First.Freq), MsgBoxStyle.OkOnly)
        'MsgBox([String].Format("ROP0_H = {0}; ROP0_V = {1}", ROP0query.First.Power_H, ROP0query.First.Power_V), MsgBoxStyle.OkOnly)

        'Create table of non zero records

        Dim NZtable = From rw In table _
                      Where rw.Volt > Zero_Volt1 _
                      And rw.Freq > ZERO_FREQ _
                      Select rw

        'MsgBox([String].Format("NZtable count = {0}", NZtable.Count), MsgBoxStyle.OkOnly)

        Dim ROPinfquery = From ropf In NZtable _
                          Order By ropf.Freq Ascending _
                          Select ropf

        If DB_Update(drw, ROPinfquery.First.power, "ROP_inf", dr) < 1 Then PrintError(drw.Record_ID, "ROP_inf")
        If DB_Update(drw, ROPinfquery.First.rms, "RMS_inf", dr) < 1 Then PrintError(drw.Record_ID, "RMS_inf")

        'MsgBox([String].Format("ROPinf_H = {0}; ROPinf_V = {1}; ROPinf_freq = {2}", ROPinfquery.First.Power_H, ROPinfquery.First.Power_V, ROPinfquery.First.Freq), MsgBoxStyle.OkOnly)


        Dim Peakquery = From pk In NZtable _
                         Order By pk.power Descending _
                         Select pk

        If DB_Update(drw, Peakquery.First.power, "PeakOP_Power", dr) < 1 Then PrintError(drw.Record_ID, "PeakOP_Power")
        If DB_Update(drw, Peakquery.First.rms, "PeakOP_RMS", dr) < 1 Then PrintError(drw.Record_ID, "PeakOP_RMS")
        If DB_Update(drw, Peakquery.First.Volt, "PeakOP_Volt", dr) < 1 Then PrintError(drw.Record_ID, "PeakOP_Volt")
        If DB_Update(drw, Peakquery.First.Freq, "PeakOP_Freq", dr) < 1 Then PrintError(drw.Record_ID, "PeakOP_Freq")

        'MsgBox([String].Format("PeakP = {0}; PeakF = {1}; PeakRMS = {2}", PeakHquery.First.Power_H, PeakHquery.First.Freq, PeakHquery.First.RMS_H), MsgBoxStyle.OkOnly)

        If (Peakquery.First.Freq < FreqMaxquery.First.Freq) _
        And (Peakquery.First.Freq > FreqMinquery.First.Freq) Then

            Dim PFq = (From fpk In NZtable _
                      Where fpk.Index >= Peakquery.First.Index - 1 _
                      And fpk.Index <= Peakquery.First.Index + 1 _
                      Order By fpk.Freq Ascending _
                      Select x = fpk.Freq, y = fpk.power, r = fpk.rms).Take(3)

            If PFq.Count > 2 Then

                For i = 0 To 2
                    xy(0, i) = PFq(i).x
                    xy(1, i) = PFq(i).y
                    xr(0, i) = PFq(i).x
                    xr(1, i) = PFq(i).r
                Next
                'MsgBox([String].Format("Freq0 = {0}; Freq1 = {1}; Freq2 = {2}", FindPeakHquery(0).Freq, FindPeakHquery(1).Freq, FindPeakHquery(2).Freq), MsgBoxStyle.OkOnly)

                MJNFitting.Fitting.QuadFit3(xy, coeff_y)
                MJNFitting.Fitting.QuadFit3(xr, coeff_r)

                temp = MJNFitting.Fitting.QuadPeakX(coeff_y)
                temp = coeff_r(0) * temp ^ 2 + coeff_r(1) * temp + coeff_r(2)

                If DB_Update(drw, MJNFitting.Fitting.QuadPeakY(coeff_y), "PeakOPfit_Power", dr) < 1 Then PrintError(drw.Record_ID, "PeakOPfit_Power")
                If DB_Update(drw, MJNFitting.Fitting.QuadPeakX(coeff_y), "PeakOPfit_Freq", dr) < 1 Then PrintError(drw.Record_ID, "PeakOPfit_Freq")
                If DB_Update(drw, temp, "PeakOPfit_RMS", dr) < 1 Then PrintError(drw.Record_ID, "PeakOPfit_RMS")

            End If

        End If

        Analyze_ClearOP(drw, "ClearOP25", dr, NZtable)
        Analyze_ClearOP(drw, "ClearOP20", dr, NZtable)
        Analyze_ClearOP(drw, "ClearOP15", dr, NZtable)

        Dim FCq = (From clr In NZtable _
                  Where clr.Freq <= Peakquery.First.Freq _
                  And clr.rms <= 0.25 _
                  Order By clr.Freq Descending _
                  Select clr).Take(1)

        'MsgBox([String].Format("Freq0 = {0}; Power0 = {1}; RMS0 = {2}", PeakHquery.First.Freq, PeakHquery.First.Power_H, PeakHquery.First.RMS_H), MsgBoxStyle.OkOnly)

        'MsgBox([String].Format("Clear Count = {0}", FindClearHquery.Count), MsgBoxStyle.OkOnly)

        If FCq.Count = 0 Then

        Else
            'MsgBox([String].Format("Freq0 = {0}; FreqCl = {1}; rmsCl = {2}", PeakHquery.First.Freq, FindClearHquery.First.Freq, FindClearHquery.First.RMS_H), MsgBoxStyle.OkOnly)
            If DB_Update(drw, FCq.First.power, "ClearOP_Power", dr) < 1 Then PrintError(drw.Record_ID, "ClearOP_Power")
            If DB_Update(drw, FCq.First.rms, "ClearOP_RMS", dr) < 1 Then PrintError(drw.Record_ID, "ClearOP_RMS")
            If DB_Update(drw, FCq.First.Volt, "ClearOP_Volt", dr) < 1 Then PrintError(drw.Record_ID, "ClearOP_Volt")
            If DB_Update(drw, FCq.First.Freq, "ClearOP_Freq", dr) < 1 Then PrintError(drw.Record_ID, "ClearOP_Freq")

        End If

        Diopt_Power(drw, 8 + ROP0query.First.power, "I08Dpt0", dr, NZtable)
        Diopt_Power(drw, FreqPower + ROP0query.First.power, "I10Dpt0", dr, NZtable)
        Diopt_Power(drw, 8 + ROPinfquery.First.power, "I08Dptinf", dr, NZtable)
        Diopt_Power(drw, FreqPower + ROPinfquery.First.power, "I10Dptinf", dr, NZtable)

        'MsgBox([String].Format("Freq Power = {0}", FreqPower), MsgBoxStyle.OkOnly)

        If 1 = 0 Then

            ref = 8 + ROP0query.First.power

            'Diopt_Power(rec_id, ref, "I08Dpt0", dr, NZtable)

            Dim FO8q = (From o8 In NZtable _
                        Where o8.power >= ref _
                        Order By o8.Freq Ascending _
                        Select o8).Take(1)
            If FO8q.Count > 0 Then

                If FO8q.First.Freq > FreqMinquery.First.Freq Then

                    Dim FIq = (From int In NZtable _
                                    Where int.Freq <= FO8q.First.Freq _
                                    Order By int.Freq Descending _
                                    Select int).Take(2)

                    'dr.I8Dp_Volt = FIq.First.Volt

                    If FIq.Count > 1 Then



                        If DB_Update(drw, FIq.First.Volt, "I8Dp_Volt", dr) < 1 Then PrintError(drw.Record_ID, "I8Dp_Volt")

                        temp = MJNFitting.Fitting.Interpolate(FIq(1).Freq, _
                                                      FIq(0).Freq, _
                                                      FIq(1).power, _
                                                      FIq(0).power, ref)

                        If DB_Update(drw, temp, "I8Dp_Freq", dr) < 1 Then PrintError(drw.Record_ID, "I8Dp_Freq")

                        temp = MJNFitting.Fitting.Interpolate(FIq(1).rms, _
                                                     FIq(0).rms, _
                                                     FIq(1).power, _
                                                     FIq(0).power, ref)

                        If DB_Update(drw, temp, "I8Dp_RMS", dr) < 1 Then PrintError(drw.Record_ID, "I8Dp_RMS")

                    End If

                End If

            End If

            Dim FO10q = (From o10 In NZtable _
                Where o10.power >= 10 _
                Order By o10.Freq Ascending _
                Select o10).Take(1)
            If FO10q.Count > 0 Then

                If FO10q.First.Freq > FreqMinquery.First.Freq Then

                    Dim FIq = (From int In NZtable _
                                        Where int.Freq <= FO10q.First.Freq _
                                        Order By int.Freq Descending _
                                        Select int).Take(2)
                    If FIq.Count > 1 Then


                        temp = MJNFitting.Fitting.Interpolate(FIq(1).Freq, _
                                           FIq(0).Freq, _
                                           FIq(1).power, _
                                           FIq(0).power, 10 + ROP0query.First.power)

                        If DB_Update(drw, temp, "I10Dp_Freq", dr) < 1 Then PrintError(drw.Record_ID, "I10Dp_Freq")

                        temp = MJNFitting.Fitting.Interpolate(FIq(1).rms, _
                                                     FIq(0).rms, _
                                                     FIq(1).power, _
                                                     FIq(0).power, 10 + ROP0query.First.power)

                        If DB_Update(drw, temp, "I10Dp_RMS", dr) < 1 Then PrintError(drw.Record_ID, "I10Dp_RMS")

                    End If

                End If

            End If

        End If
        'dr.Grade = 0

    End Sub


    Public Function DB_Update_Head(ByVal RecordID As String)


        Dim cur = (From wr In DB_Con.ALF_Datas _
                  Where wr.Record_ID Is RecordID _
                  Select wr).ToList()(0)

        Dim row, col As Integer
        Dim temp, retval As Integer
        Dim colStr, rowStr As String

        retval = 0

        Select Case Strings.Left(cur.PartNumber, 2)
            Case "FL"
                colStr = Strings.Right(Strings.Left(cur.PartNumber, 12), 2)
                rowStr = Strings.Right(Strings.Left(cur.PartNumber, 14), 2)
                col = Asc(Strings.Right(Strings.Left(colStr, 2), 1)) - 64
                row = Integer.Parse(rowStr)
            Case "PL"
                colStr = Strings.Right(Strings.Left(cur.PartNumber, 12), 2)
                rowStr = Strings.Right(Strings.Left(cur.PartNumber, 14), 2)
                col = Asc(Strings.Right(Strings.Left(colStr, 2), 1)) - 64
                row = Integer.Parse(rowStr)
                'row = Asc(Strings.Right(Strings.Left(row, 16), 1)) - 48
            Case Else
                colStr = Strings.Right(Strings.Left(cur.PartNumber, 15), 1)
                rowStr = Strings.Right(Strings.Left(cur.PartNumber, 16), 1)
                col = Asc(colStr) - 64
                row = Asc(rowStr) - 48
        End Select

        'col = Asc(Strings.Right(Strings.Left(cur.PartNumber, 15), 1)) - 64
        'row = Asc(Strings.Right(Strings.Left(cur.PartNumber, 16), 1)) - 48

        If col < 0 Then col = 0

        Dim CheckAnDBquery = From cadb In DB_Con.SH_Analysis _
                             Where cadb.Record_ID Is RecordID _
                             Select cadb

        temp = CheckAnDBquery.Count

        If CheckAnDBquery.Count < 1 Then
            Dim cts As New SH_Analysi With { _
            .Record_ID = RecordID, _
            .PartNumber = cur.PartNumber, _
            .ArrayNumber = cur.ArrayNumber, _
            .StartDateTime = cur.StartDateTime, _
            .Description = cur.Description, _
            .PassFailCode = cur.PassFailCode, _
            .AnalysisDateTime = DateTime.Now, _
            .AnalysisVersion = VERSION, _
            .Col = col, _
            .Row = row}

            DB_Con.SH_Analysis.InsertOnSubmit(cts)

            Try
                DB_Con.SubmitChanges()
                retval = 1
            Catch
                ' Handle exception.
                retval = -1
            End Try

        Else
            Dim crs = (From cr In DB_Con.SH_Analysis _
                        Where cr.Record_ID Is RecordID _
                        Select cr).ToList()(0)

            With crs
                .PartNumber = cur.PartNumber
                .ArrayNumber = cur.ArrayNumber
                .StartDateTime = cur.StartDateTime
                .Description = cur.Description
                .PassFailCode = cur.PassFailCode
                .AnalysisDateTime = DateTime.Now
                .AnalysisVersion = VERSION
                .Col = col
                .Row = row
            End With

            Try
                DB_Con.SubmitChanges()
                retval = 1
            Catch
                MsgBox("Update Didn't Work!", MsgBoxStyle.OkOnly)
                retval = -1
            End Try

        End If
        Return retval
    End Function


    Public Function DB_Update_Head_MES(ByVal RecordID As String)


        Dim cur = (From wr In DB_Con.ALF_Data_MEs _
                  Where wr.Record_ID Is RecordID _
                  Select wr).ToList()(0)

        Dim row, col As Integer
        Dim temp, retval As Integer
        Dim colStr, rowStr As String

        retval = 0

        Select Case Strings.Left(cur.PartNumber, 2)
            Case "FL"
                colStr = Strings.Right(Strings.Left(cur.PartNumber, 12), 2)
                rowStr = Strings.Right(Strings.Left(cur.PartNumber, 14), 2)
                col = Asc(Strings.Right(Strings.Left(colStr, 2), 1)) - 64
                row = Integer.Parse(rowStr)
                'row = Asc(Strings.Right(Strings.Left(row, 16), 1)) - 48
            Case Else
                colStr = Strings.Right(Strings.Left(cur.PartNumber, 15), 1)
                rowStr = Strings.Right(Strings.Left(cur.PartNumber, 16), 1)
                col = Asc(colStr) - 64
                row = Asc(rowStr) - 48
        End Select

        'col = Asc(Strings.Right(Strings.Left(cur.PartNumber, 15), 1)) - 64
        'row = Asc(Strings.Right(Strings.Left(cur.PartNumber, 16), 1)) - 48

        If col < 0 Then col = 0

        Dim CheckAnDBquery = From cadb In DB_Con.SH_Analysis _
                             Where cadb.Record_ID Is RecordID _
                             Select cadb

        temp = CheckAnDBquery.Count

        If CheckAnDBquery.Count < 1 Then
            Dim cts As New SH_Analysi With { _
            .Record_ID = RecordID, _
            .PartNumber = cur.PartNumber, _
            .ArrayNumber = cur.ArrayNumber, _
            .StartDateTime = cur.StartDateTime, _
            .Description = cur.Description, _
            .PassFailCode = cur.PassFailCode, _
            .AnalysisDateTime = DateTime.Now, _
            .AnalysisVersion = VERSION, _
            .Col = col, _
            .Row = row}

            DB_Con.SH_Analysis.InsertOnSubmit(cts)

            Try
                DB_Con.SubmitChanges()
                retval = 1
            Catch
                ' Handle exception.
                retval = -1
            End Try

        Else
            Dim crs = (From cr In DB_Con.SH_Analysis _
                        Where cr.Record_ID Is RecordID _
                        Select cr).ToList()(0)

            With crs
                .PartNumber = cur.PartNumber
                .ArrayNumber = cur.ArrayNumber
                .StartDateTime = cur.StartDateTime
                .Description = cur.Description
                .PassFailCode = cur.PassFailCode
                .AnalysisDateTime = DateTime.Now
                .AnalysisVersion = VERSION
                .Col = col
                .Row = row
            End With

            Try
                DB_Con.SubmitChanges()
                retval = 1
            Catch
                MsgBox("Update Didn't Work!", MsgBoxStyle.OkOnly)
                retval = -1
            End Try

        End If
        Return retval
    End Function

    Private Function DB_Update(ByRef drw As DataSet2.ResultsTableRow, _
                              ByVal value As Double, _
                              ByVal Field As String, _
                              ByVal Pol As Integer) As Integer

        Dim temp As Double
        Dim retval As Integer = 1

        If Not Double.IsNaN(value) Then
            temp = AsDoub(value)

            With drw
                Select Case Pol
                    Case 0
                        Select Case Field
                            Case "ROP_0"
                                .ROP_0_avg = temp
                            Case "RMS_0"
                                .RMS_0_avg = temp
                            Case "ROP_inf"
                                .ROP_inf_avg = temp
                            Case "RMS_inf"
                                .RMS_inf_avg = temp

                            Case "ClearOP_Power"
                                .ClearOP_Power_avg = temp
                            Case "ClearOP_RMS"
                                .ClearOP_RMS_avg = temp
                            Case "ClearOP_Volt"
                                .ClearOP_Volt_avg = temp
                            Case "ClearOP_Freq"
                                .ClearOP_Freq_avg = temp

                            Case "I8Dp_Volt"
                                .I8Dp_Volt_avg = temp
                            Case "I10Dp_Freq"
                                .I10Dp_Freq_avg = temp
                            Case "I10Dp_RMS"
                                .I10Dp_RMS_avg = temp
                            Case "I8Dp_Freq"
                                .I8Dp_Freq_avg = temp
                            Case "I8Dp_RMS"
                                .I8Dp_RMS_avg = temp

                            Case "PeakOP_Power"
                                .PeakOP_Power_avg = temp
                            Case "PeakOP_RMS"
                                .PeakOP_RMS_avg = temp
                            Case "PeakOP_Volt"
                                .PeakOP_Volt_avg = temp
                            Case "PeakOP_Freq"
                                .PeakOP_Freq_avg = temp

                            Case "PeakOPfit_Power"
                                .PeakOPfit_Power_avg = temp
                            Case "PeakOPfit_RMS"
                                .PeakOPfit_RMS_avg = temp
                            Case "PeakOPfit_Freq"
                                .PeakOPfit_Freq_avg = temp

                            Case "Pol_Off"
                                .Polarization_Offset = temp
                            Case "Max_Pol_Off"
                                .Max_Delta_Polar_Power_12D = temp

                            Case "I08Dpt0_Volt"
                                .I08Dpt0_Volt_avg = temp
                                .I8Dp_Volt_avg = temp
                            Case "I08Dpt0_Freq"
                                .I08Dpt0_Freq_avg = temp
                                .I8Dp_Freq_avg = temp
                            Case "I08Dpt0_RMS"
                                .I08Dpt0_RMS_avg = temp
                                .I8Dp_RMS_avg = temp

                            Case "I10Dpt0_Volt"
                                '.I10Dpt0_Volt_avg = temp
                            Case "I10Dpt0_Freq"
                                .I10Dpt0_Freq_avg = temp
                                .I10Dp_Freq_avg = temp
                            Case "I10Dpt0_RMS"
                                .I10Dpt0_RMS_avg = temp
                                .I10Dp_RMS_avg = temp

                            Case "I08Dptinf_Volt"
                                '.I08Dptinf_Volt_avg = temp
                            Case "I08Dptinf_Freq"
                                .I08Dptinf_Freq_avg = temp
                            Case "I08Dptinf_RMS"
                                .I08Dptinf_RMS_avg = temp

                            Case "I10Dptinf_Volt"
                                '.I10Dptinf_Volt_avg = temp
                            Case "I10Dptinf_Freq"
                                .I10Dptinf_Freq_avg = temp
                            Case "I10Dptinf_RMS"
                                .I10Dptinf_RMS_avg = temp

                            Case "ClearOP25_Power"
                                .ClearOP25_Power_avg = temp
                            Case "ClearOP25_RMS"
                                .ClearOP25_RMS_avg = temp
                            Case "ClearOP25_Volt"
                                .ClearOP25_Volt_avg = temp
                            Case "ClearOP25_Freq"
                                .ClearOP25_Freq_avg = temp

                            Case "ClearOP20_Power"
                                .ClearOP20_Power_avg = temp
                            Case "ClearOP20_RMS"
                                .ClearOP20_RMS_avg = temp
                            Case "ClearOP20_Volt"
                                .ClearOP20_Volt_avg = temp
                            Case "ClearOP20_Freq"
                                .ClearOP20_Freq_avg = temp

                            Case "ClearOP15_Power"
                                .ClearOP15_Power_avg = temp
                            Case "ClearOP15_RMS"
                                .ClearOP15_RMS_avg = temp
                            Case "ClearOP15_Volt"
                                .ClearOP15_Volt_avg = temp
                            Case "ClearOP15_Freq"
                                .ClearOP15_Freq_avg = temp

                            Case "Future_1"
                                .Future_1 = temp
                            Case "Future_2"
                                .Future_2 = temp
                            Case "Future_3"
                                .Future_3 = temp
                        End Select
                    Case 1
                        Select Case Field
                            Case "ROP_0"
                                .ROP_0_H = temp
                            Case "RMS_0"
                                .RMS_0_H = temp
                            Case "ROP_inf"
                                .ROP_inf_H = temp
                            Case "RMS_inf"
                                .RMS_inf_H = temp

                            Case "ClearOP_Power"
                                .ClearOP_Power_H = temp
                            Case "ClearOP_RMS"
                                .ClearOP_RMS_H = temp
                            Case "ClearOP_Volt"
                                .ClearOP_Volt_H = temp
                            Case "ClearOP_Freq"
                                .ClearOP_Freq_H = temp

                            Case "I8Dp_Volt"
                                .I8Dp_Volt_H = temp
                            Case "I10Dp_Freq"
                                .I10Dp_Freq_H = temp
                            Case "I10Dp_RMS"
                                .I10Dp_RMS_H = temp
                            Case "I8Dp_Freq"
                                .I8Dp_Freq_H = temp
                            Case "I8Dp_RMS"
                                .I8Dp_RMS_H = temp

                            Case "PeakOP_Power"
                                .PeakOP_Power_H = temp
                            Case "PeakOP_RMS"
                                .PeakOP_RMS_H = temp
                            Case "PeakOP_Volt"
                                .PeakOP_Volt_H = temp
                            Case "PeakOP_Freq"
                                .PeakOP_Freq_H = temp

                            Case "PeakOPfit_Power"
                                .PeakOPfit_Power_H = temp
                            Case "PeakOPfit_RMS"
                                .PeakOPfit_RMS_H = temp
                            Case "PeakOPfit_Freq"
                                .PeakOPfit_Freq_H = temp

                            Case "I08Dpt0_Volt"
                                .I08Dpt0_Volt_H = temp
                            Case "I08Dpt0_Freq"
                                .I08Dpt0_Freq_H = temp
                            Case "I08Dpt0_RMS"
                                .I08Dpt0_RMS_H = temp

                            Case "I10Dpt0_Volt"
                                '.I10Dpt0_Volt_H = temp
                            Case "I10Dpt0_Freq"
                                .I10Dpt0_Freq_H = temp
                            Case "I10Dpt0_RMS"
                                .I10Dpt0_RMS_H = temp

                            Case "I08Dptinf_Volt"
                                '.I08Dptinf_Volt_H = temp
                            Case "I08Dptinf_Freq"
                                .I08Dptinf_Freq_H = temp
                            Case "I08Dptinf_RMS"
                                .I08Dptinf_RMS_H = temp

                            Case "I10Dptinf_Volt"
                                '.I10Dptinf_Volt_H = temp
                            Case "I10Dptinf_Freq"
                                .I10Dptinf_Freq_H = temp
                            Case "I10Dptinf_RMS"
                                .I10Dptinf_RMS_H = temp

                            Case "ClearOP25_Power"
                                .ClearOP25_Power_H = temp
                            Case "ClearOP25_RMS"
                                .ClearOP25_RMS_H = temp
                            Case "ClearOP25_Volt"
                                .ClearOP25_Volt_H = temp
                            Case "ClearOP25_Freq"
                                .ClearOP25_Freq_H = temp

                            Case "ClearOP20_Power"
                                .ClearOP20_Power_H = temp
                            Case "ClearOP20_RMS"
                                .ClearOP20_RMS_H = temp
                            Case "ClearOP20_Volt"
                                .ClearOP20_Volt_H = temp
                            Case "ClearOP20_Freq"
                                .ClearOP20_Freq_H = temp

                            Case "ClearOP15_Power"
                                .ClearOP15_Power_H = temp
                            Case "ClearOP15_RMS"
                                .ClearOP15_RMS_H = temp
                            Case "ClearOP15_Volt"
                                .ClearOP15_Volt_H = temp
                            Case "ClearOP15_Freq"
                                .ClearOP15_Freq_H = temp
                        End Select
                    Case 2
                        Select Case Field
                            Case "ROP_0"
                                .ROP_0_V = temp
                            Case "RMS_0"
                                .RMS_0_V = temp
                            Case "ROP_inf"
                                .ROP_inf_V = temp
                            Case "RMS_inf"
                                .RMS_inf_V = temp

                            Case "ClearOP_Power"
                                .ClearOP_Power_V = temp
                            Case "ClearOP_RMS"
                                .ClearOP_RMS_V = temp
                            Case "ClearOP_Volt"
                                .ClearOP_Volt_V = temp
                            Case "ClearOP_Freq"
                                .ClearOP_Freq_V = temp

                            Case "I8Dp_Volt"
                                .I8Dp_Volt_V = temp

                            Case "I10Dp_Freq"
                                .I10Dp_Freq_V = temp
                            Case "I10Dp_RMS"
                                .I10Dp_RMS_V = temp

                            Case "I8Dp_Freq"
                                .I8Dp_Freq_V = temp
                            Case "I8Dp_RMS"
                                .I8Dp_RMS_V = temp

                            Case "PeakOP_Power"
                                .PeakOP_Power_V = temp
                            Case "PeakOP_RMS"
                                .PeakOP_RMS_V = temp
                            Case "PeakOP_Volt"
                                .PeakOP_Volt_V = temp
                            Case "PeakOP_Freq"
                                .PeakOP_Freq_V = temp

                            Case "PeakOPfit_Power"
                                .PeakOPfit_Power_V = temp
                            Case "PeakOPfit_RMS"
                                .PeakOPfit_RMS_V = temp
                            Case "PeakOPfit_Freq"
                                .PeakOPfit_Freq_V = temp

                            Case "I08Dpt0_Volt"
                                .I08Dpt0_Volt_V = temp
                            Case "I08Dpt0_Freq"
                                .I08Dpt0_Freq_V = temp
                            Case "I08Dpt0_RMS"
                                .I08Dpt0_RMS_V = temp

                            Case "I10Dpt0_Volt"
                                '.I10Dpt0_Volt_V = temp
                            Case "I10Dpt0_Freq"
                                .I10Dpt0_Freq_V = temp
                            Case "I10Dpt0_RMS"
                                .I10Dpt0_RMS_V = temp

                            Case "I08Dptinf_Volt"
                                '.I08Dptinf_Volt_V = temp
                            Case "I08Dptinf_Freq"
                                .I08Dptinf_Freq_V = temp
                            Case "I08Dptinf_RMS"
                                .I08Dptinf_RMS_V = temp

                            Case "I10Dptinf_Volt"
                                '.I10Dptinf_Volt_V = temp
                            Case "I10Dptinf_Freq"
                                .I10Dptinf_Freq_V = temp
                            Case "I10Dptinf_RMS"
                                .I10Dptinf_RMS_V = temp

                            Case "ClearOP25_Power"
                                .ClearOP25_Power_V = temp
                            Case "ClearOP25_RMS"
                                .ClearOP25_RMS_V = temp
                            Case "ClearOP25_Volt"
                                .ClearOP25_Volt_V = temp
                            Case "ClearOP25_Freq"
                                .ClearOP25_Freq_V = temp

                            Case "ClearOP20_Power"
                                .ClearOP20_Power_V = temp
                            Case "ClearOP20_RMS"
                                .ClearOP20_RMS_V = temp
                            Case "ClearOP20_Volt"
                                .ClearOP20_Volt_V = temp
                            Case "ClearOP20_Freq"
                                .ClearOP20_Freq_V = temp

                            Case "ClearOP15_Power"
                                .ClearOP15_Power_V = temp
                            Case "ClearOP15_RMS"
                                .ClearOP15_RMS_V = temp
                            Case "ClearOP15_Volt"
                                .ClearOP15_Volt_V = temp
                            Case "ClearOP15_Freq"
                                .ClearOP15_Freq_V = temp
                        End Select
                End Select

            End With
        Else
            'retval = -1
        End If
        Return retval

    End Function

    Public Function DB_Update2(ByVal drw As DataSet2.ResultsTableRow) As Integer

        Dim retval As Integer = 0

        Dim CheckAnDBquery = From cadb In DB_Con.SH_Analysis _
                             Where cadb.Record_ID Is drw.Record_ID _
                             Select cadb

        If CheckAnDBquery.Count > 0 Then


            Dim crs = (From cr In DB_Con.SH_Analysis _
                        Where cr.Record_ID Is drw.Record_ID _
                        Select cr).ToList()(0)
            With crs

                If Not Double.IsNaN(drw.ROP_0_avg) Then .ROP_0_avg = AsDoub(drw.ROP_0_avg)
                If Not Double.IsNaN(drw.RMS_0_avg) Then .RMS_0_avg = AsDoub(drw.RMS_0_avg)
                If Not Double.IsNaN(drw.ROP_inf_avg) Then .ROP_inf_avg = AsDoub(drw.ROP_inf_avg)
                If Not Double.IsNaN(drw.RMS_inf_avg) Then .RMS_inf_avg = AsDoub(drw.RMS_inf_avg)

                If Not Double.IsNaN(drw.ClearOP_Power_avg) Then .ClearOP_Power_avg = AsDoub(drw.ClearOP_Power_avg)
                If Not Double.IsNaN(drw.ClearOP_RMS_avg) Then .ClearOP_RMS_avg = AsDoub(drw.ClearOP_RMS_avg)
                If Not Double.IsNaN(drw.ClearOP_Volt_avg) Then .ClearOP_Volt_avg = AsDoub(drw.ClearOP_Volt_avg)
                If Not Double.IsNaN(drw.ClearOP_Freq_avg) Then .ClearOP_Freq_avg = AsDoub(drw.ClearOP_Freq_avg)

                If Not Double.IsNaN(drw.I8Dp_Volt_avg) Then .I8Dp_Volt_avg = AsDoub(drw.I8Dp_Volt_avg)
                If Not Double.IsNaN(drw.I10Dp_Freq_avg) Then .I10Dp_Freq_avg = AsDoub(drw.I10Dp_Freq_avg)
                If Not Double.IsNaN(drw.I10Dp_RMS_avg) Then .I10Dp_RMS_avg = AsDoub(drw.I10Dp_RMS_avg)
                If Not Double.IsNaN(drw.I8Dp_Freq_avg) Then .I8Dp_Freq_avg = AsDoub(drw.I8Dp_Freq_avg)
                If Not Double.IsNaN(drw.I8Dp_RMS_avg) Then .I8Dp_RMS_avg = AsDoub(drw.I8Dp_RMS_avg)

                If Not Double.IsNaN(drw.PeakOP_Power_avg) Then .PeakOP_Power_avg = AsDoub(drw.PeakOP_Power_avg)
                If Not Double.IsNaN(drw.PeakOP_RMS_avg) Then .PeakOP_RMS_avg = AsDoub(drw.PeakOP_RMS_avg)
                If Not Double.IsNaN(drw.PeakOP_Volt_avg) Then .PeakOP_Volt_avg = AsDoub(drw.PeakOP_Volt_avg)
                If Not Double.IsNaN(drw.PeakOP_Freq_avg) Then .PeakOP_Freq_avg = AsDoub(drw.PeakOP_Freq_avg)

                If Not Double.IsNaN(drw.PeakOPfit_Power_avg) Then .PeakOPfit_Power_avg = AsDoub(drw.PeakOPfit_Power_avg)
                If Not Double.IsNaN(drw.PeakOPfit_RMS_avg) Then .PeakOPfit_RMS_avg = AsDoub(drw.PeakOPfit_RMS_avg)
                If Not Double.IsNaN(drw.PeakOPfit_Freq_avg) Then .PeakOPfit_Freq_avg = AsDoub(drw.PeakOPfit_Freq_avg)

                If Not Double.IsNaN(drw.Polarization_Offset) Then .Polarization_Offset = AsDoub(drw.Polarization_Offset)
                If Not Double.IsNaN(drw.Max_Delta_Polar_Power_12D) Then .Max_Delta_Polar_Power_12D = AsDoub(drw.Max_Delta_Polar_Power_12D)

                If Not Double.IsNaN(drw.I08Dpt0_Volt_avg) Then .I08Dpt0_Volt_avg = AsDoub(drw.I08Dpt0_Volt_avg)
                If Not Double.IsNaN(drw.I08Dpt0_Freq_avg) Then .I08Dpt0_Freq_avg = AsDoub(drw.I08Dpt0_Freq_avg)
                If Not Double.IsNaN(drw.I08Dpt0_RMS_avg) Then .I08Dpt0_RMS_avg = AsDoub(drw.I08Dpt0_RMS_avg)

                If Not Double.IsNaN(drw.I10Dpt0_Freq_avg) Then .I10Dpt0_Freq_avg = AsDoub(drw.I10Dpt0_Freq_avg)
                If Not Double.IsNaN(drw.I10Dpt0_RMS_avg) Then .I10Dpt0_RMS_avg = AsDoub(drw.I10Dpt0_RMS_avg)

                If Not Double.IsNaN(drw.I08Dptinf_Freq_avg) Then .I08Dptinf_Freq_avg = AsDoub(drw.I08Dptinf_Freq_avg)
                If Not Double.IsNaN(drw.I08Dptinf_RMS_avg) Then .I08Dptinf_RMS_avg = AsDoub(drw.I08Dptinf_RMS_avg)

                If Not Double.IsNaN(drw.I10Dptinf_Freq_avg) Then .I10Dptinf_Freq_avg = AsDoub(drw.I10Dptinf_Freq_avg)
                If Not Double.IsNaN(drw.I10Dptinf_RMS_avg) Then .I10Dptinf_RMS_avg = AsDoub(drw.I10Dptinf_RMS_avg)

                If Not Double.IsNaN(drw.ClearOP25_Power_avg) Then .ClearOP25_Power_avg = AsDoub(drw.ClearOP25_Power_avg)
                If Not Double.IsNaN(drw.ClearOP25_RMS_avg) Then .ClearOP25_RMS_avg = AsDoub(drw.ClearOP25_RMS_avg)
                If Not Double.IsNaN(drw.ClearOP25_Volt_avg) Then .ClearOP25_Volt_avg = AsDoub(drw.ClearOP25_Volt_avg)
                If Not Double.IsNaN(drw.ClearOP25_Freq_avg) Then .ClearOP25_Freq_avg = AsDoub(drw.ClearOP25_Freq_avg)

                If Not Double.IsNaN(drw.ClearOP20_Power_avg) Then .ClearOP20_Power_avg = AsDoub(drw.ClearOP20_Power_avg)
                If Not Double.IsNaN(drw.ClearOP20_RMS_avg) Then .ClearOP20_RMS_avg = AsDoub(drw.ClearOP20_RMS_avg)
                If Not Double.IsNaN(drw.ClearOP20_Volt_avg) Then .ClearOP20_Volt_avg = AsDoub(drw.ClearOP20_Volt_avg)
                If Not Double.IsNaN(drw.ClearOP20_Freq_avg) Then .ClearOP20_Freq_avg = AsDoub(drw.ClearOP20_Freq_avg)

                If Not Double.IsNaN(drw.ClearOP15_Power_avg) Then .ClearOP15_Power_avg = AsDoub(drw.ClearOP15_Power_avg)
                If Not Double.IsNaN(drw.ClearOP15_RMS_avg) Then .ClearOP15_RMS_avg = AsDoub(drw.ClearOP15_RMS_avg)
                If Not Double.IsNaN(drw.ClearOP15_Volt_avg) Then .ClearOP15_Volt_avg = AsDoub(drw.ClearOP15_Volt_avg)
                If Not Double.IsNaN(drw.ClearOP15_Freq_avg) Then .ClearOP15_Freq_avg = AsDoub(drw.ClearOP15_Freq_avg)

                If Not Double.IsNaN(drw.ROP_0_H) Then .ROP_0_H = AsDoub(drw.ROP_0_H)
                If Not Double.IsNaN(drw.RMS_0_H) Then .RMS_0_H = AsDoub(drw.RMS_0_H)
                If Not Double.IsNaN(drw.ROP_inf_H) Then .ROP_inf_H = AsDoub(drw.ROP_inf_H)
                If Not Double.IsNaN(drw.RMS_inf_H) Then .RMS_inf_H = AsDoub(drw.RMS_inf_H)

                If Not Double.IsNaN(drw.ClearOP_Power_H) Then .ClearOP_Power_H = AsDoub(drw.ClearOP_Power_H)
                If Not Double.IsNaN(drw.ClearOP_RMS_H) Then .ClearOP_RMS_H = AsDoub(drw.ClearOP_RMS_H)
                If Not Double.IsNaN(drw.ClearOP_Volt_H) Then .ClearOP_Volt_H = AsDoub(drw.ClearOP_Volt_H)
                If Not Double.IsNaN(drw.ClearOP_Freq_H) Then .ClearOP_Freq_H = AsDoub(drw.ClearOP_Freq_H)

                If Not Double.IsNaN(drw.I8Dp_Volt_H) Then .I8Dp_Volt_H = AsDoub(drw.I8Dp_Volt_H)
                If Not Double.IsNaN(drw.I10Dp_Freq_H) Then .I10Dp_Freq_H = AsDoub(drw.I10Dp_Freq_H)
                If Not Double.IsNaN(drw.I10Dp_RMS_H) Then .I10Dp_RMS_H = AsDoub(drw.I10Dp_RMS_H)
                If Not Double.IsNaN(drw.I8Dp_Freq_H) Then .I8Dp_Freq_H = AsDoub(drw.I8Dp_Freq_H)
                If Not Double.IsNaN(drw.I8Dp_RMS_H) Then .I8Dp_RMS_H = AsDoub(drw.I8Dp_RMS_H)

                If Not Double.IsNaN(drw.PeakOP_Power_H) Then .PeakOP_Power_H = AsDoub(drw.PeakOP_Power_H)
                If Not Double.IsNaN(drw.PeakOP_RMS_H) Then .PeakOP_RMS_H = AsDoub(drw.PeakOP_RMS_H)
                If Not Double.IsNaN(drw.PeakOP_Volt_H) Then .PeakOP_Volt_H = AsDoub(drw.PeakOP_Volt_H)
                If Not Double.IsNaN(drw.PeakOP_Freq_H) Then .PeakOP_Freq_H = AsDoub(drw.PeakOP_Freq_H)

                If Not Double.IsNaN(drw.PeakOPfit_Power_H) Then .PeakOPfit_Power_H = AsDoub(drw.PeakOPfit_Power_H)
                If Not Double.IsNaN(drw.PeakOPfit_RMS_H) Then .PeakOPfit_RMS_H = AsDoub(drw.PeakOPfit_RMS_H)
                If Not Double.IsNaN(drw.PeakOPfit_Freq_H) Then .PeakOPfit_Freq_H = AsDoub(drw.PeakOPfit_Freq_H)

                If Not Double.IsNaN(drw.I08Dpt0_Volt_H) Then .I08Dpt0_Volt_H = AsDoub(drw.I08Dpt0_Volt_H)
                If Not Double.IsNaN(drw.I08Dpt0_Freq_H) Then .I08Dpt0_Freq_H = AsDoub(drw.I08Dpt0_Freq_H)
                If Not Double.IsNaN(drw.I08Dpt0_RMS_H) Then .I08Dpt0_RMS_H = AsDoub(drw.I08Dpt0_RMS_H)

                If Not Double.IsNaN(drw.I10Dpt0_Freq_H) Then .I10Dpt0_Freq_H = AsDoub(drw.I10Dpt0_Freq_H)
                If Not Double.IsNaN(drw.I10Dpt0_RMS_H) Then .I10Dpt0_RMS_H = AsDoub(drw.I10Dpt0_RMS_H)

                If Not Double.IsNaN(drw.I08Dptinf_Freq_H) Then .I08Dptinf_Freq_H = AsDoub(drw.I08Dptinf_Freq_H)
                If Not Double.IsNaN(drw.I08Dptinf_RMS_H) Then .I08Dptinf_RMS_H = AsDoub(drw.I08Dptinf_RMS_H)

                If Not Double.IsNaN(drw.I10Dptinf_Freq_H) Then .I10Dptinf_Freq_H = AsDoub(drw.I10Dptinf_Freq_H)
                If Not Double.IsNaN(drw.I10Dptinf_RMS_H) Then .I10Dptinf_RMS_H = AsDoub(drw.I10Dptinf_RMS_H)

                If Not Double.IsNaN(drw.ClearOP25_Power_H) Then .ClearOP25_Power_H = AsDoub(drw.ClearOP25_Power_H)
                If Not Double.IsNaN(drw.ClearOP25_RMS_H) Then .ClearOP25_RMS_H = AsDoub(drw.ClearOP25_RMS_H)
                If Not Double.IsNaN(drw.ClearOP25_Volt_H) Then .ClearOP25_Volt_H = AsDoub(drw.ClearOP25_Volt_H)
                If Not Double.IsNaN(drw.ClearOP25_Freq_H) Then .ClearOP25_Freq_H = AsDoub(drw.ClearOP25_Freq_H)

                If Not Double.IsNaN(drw.ClearOP20_Power_H) Then .ClearOP20_Power_H = AsDoub(drw.ClearOP20_Power_H)
                If Not Double.IsNaN(drw.ClearOP20_RMS_H) Then .ClearOP20_RMS_H = AsDoub(drw.ClearOP20_RMS_H)
                If Not Double.IsNaN(drw.ClearOP20_Volt_H) Then .ClearOP20_Volt_H = AsDoub(drw.ClearOP20_Volt_H)
                If Not Double.IsNaN(drw.ClearOP20_Freq_H) Then .ClearOP20_Freq_H = AsDoub(drw.ClearOP20_Freq_H)

                If Not Double.IsNaN(drw.ClearOP15_Power_H) Then .ClearOP15_Power_H = AsDoub(drw.ClearOP15_Power_H)
                If Not Double.IsNaN(drw.ClearOP15_RMS_H) Then .ClearOP15_RMS_H = AsDoub(drw.ClearOP15_RMS_H)
                If Not Double.IsNaN(drw.ClearOP15_Volt_H) Then .ClearOP15_Volt_H = AsDoub(drw.ClearOP15_Volt_H)
                If Not Double.IsNaN(drw.ClearOP15_Freq_H) Then .ClearOP15_Freq_H = AsDoub(drw.ClearOP15_Freq_H)

                If Not Double.IsNaN(drw.ROP_0_V) Then .ROP_0_V = AsDoub(drw.ROP_0_V)
                If Not Double.IsNaN(drw.RMS_0_V) Then .RMS_0_V = AsDoub(drw.RMS_0_V)
                If Not Double.IsNaN(drw.ROP_inf_V) Then .ROP_inf_V = AsDoub(drw.ROP_inf_V)
                If Not Double.IsNaN(drw.RMS_inf_V) Then .RMS_inf_V = AsDoub(drw.RMS_inf_V)

                If Not Double.IsNaN(drw.ClearOP_Power_V) Then .ClearOP_Power_V = AsDoub(drw.ClearOP_Power_V)
                If Not Double.IsNaN(drw.ClearOP_RMS_V) Then .ClearOP_RMS_V = AsDoub(drw.ClearOP_RMS_V)
                If Not Double.IsNaN(drw.ClearOP_Volt_V) Then .ClearOP_Volt_V = AsDoub(drw.ClearOP_Volt_V)
                If Not Double.IsNaN(drw.ClearOP_Freq_V) Then .ClearOP_Freq_V = AsDoub(drw.ClearOP_Freq_V)

                If Not Double.IsNaN(drw.I8Dp_Volt_V) Then .I8Dp_Volt_V = AsDoub(drw.I8Dp_Volt_V)

                If Not Double.IsNaN(drw.I10Dp_Freq_V) Then .I10Dp_Freq_V = AsDoub(drw.I10Dp_Freq_V)
                If Not Double.IsNaN(drw.I10Dp_RMS_V) Then .I10Dp_RMS_V = AsDoub(drw.I10Dp_RMS_V)

                If Not Double.IsNaN(drw.I8Dp_Freq_V) Then .I8Dp_Freq_V = AsDoub(drw.I8Dp_Freq_V)
                If Not Double.IsNaN(drw.I8Dp_RMS_V) Then .I8Dp_RMS_V = AsDoub(drw.I8Dp_RMS_V)

                If Not Double.IsNaN(drw.PeakOP_Power_V) Then .PeakOP_Power_V = AsDoub(drw.PeakOP_Power_V)
                If Not Double.IsNaN(drw.PeakOP_RMS_V) Then .PeakOP_RMS_V = AsDoub(drw.PeakOP_RMS_V)
                If Not Double.IsNaN(drw.PeakOP_Volt_V) Then .PeakOP_Volt_V = AsDoub(drw.PeakOP_Volt_V)
                If Not Double.IsNaN(drw.PeakOP_Freq_V) Then .PeakOP_Freq_V = AsDoub(drw.PeakOP_Freq_V)

                If Not Double.IsNaN(drw.PeakOPfit_Power_V) Then .PeakOPfit_Power_V = AsDoub(drw.PeakOPfit_Power_V)
                If Not Double.IsNaN(drw.PeakOPfit_RMS_V) Then .PeakOPfit_RMS_V = AsDoub(drw.PeakOPfit_RMS_V)
                If Not Double.IsNaN(drw.PeakOPfit_Freq_V) Then .PeakOPfit_Freq_V = AsDoub(drw.PeakOPfit_Freq_V)

                If Not Double.IsNaN(drw.I08Dpt0_Volt_V) Then .I08Dpt0_Volt_V = AsDoub(drw.I08Dpt0_Volt_V)
                If Not Double.IsNaN(drw.I08Dpt0_Freq_V) Then .I08Dpt0_Freq_V = AsDoub(drw.I08Dpt0_Freq_V)
                If Not Double.IsNaN(drw.I08Dpt0_RMS_V) Then .I08Dpt0_RMS_V = AsDoub(drw.I08Dpt0_RMS_V)

                If Not Double.IsNaN(drw.I10Dpt0_Freq_V) Then .I10Dpt0_Freq_V = AsDoub(drw.I10Dpt0_Freq_V)
                If Not Double.IsNaN(drw.I10Dpt0_RMS_V) Then .I10Dpt0_RMS_V = AsDoub(drw.I10Dpt0_RMS_V)

                If Not Double.IsNaN(drw.I08Dptinf_Freq_V) Then .I08Dptinf_Freq_V = AsDoub(drw.I08Dptinf_Freq_V)
                If Not Double.IsNaN(drw.I08Dptinf_RMS_V) Then .I08Dptinf_RMS_V = AsDoub(drw.I08Dptinf_RMS_V)

                If Not Double.IsNaN(drw.I10Dptinf_Freq_V) Then .I10Dptinf_Freq_V = AsDoub(drw.I10Dptinf_Freq_V)
                If Not Double.IsNaN(drw.I10Dptinf_RMS_V) Then .I10Dptinf_RMS_V = AsDoub(drw.I10Dptinf_RMS_V)

                If Not Double.IsNaN(drw.ClearOP25_Power_V) Then .ClearOP25_Power_V = AsDoub(drw.ClearOP25_Power_V)
                If Not Double.IsNaN(drw.ClearOP25_RMS_V) Then .ClearOP25_RMS_V = AsDoub(drw.ClearOP25_RMS_V)
                If Not Double.IsNaN(drw.ClearOP25_Volt_V) Then .ClearOP25_Volt_V = AsDoub(drw.ClearOP25_Volt_V)
                If Not Double.IsNaN(drw.ClearOP25_Freq_V) Then .ClearOP25_Freq_V = AsDoub(drw.ClearOP25_Freq_V)

                If Not Double.IsNaN(drw.ClearOP20_Power_V) Then .ClearOP20_Power_V = AsDoub(drw.ClearOP20_Power_V)
                If Not Double.IsNaN(drw.ClearOP20_RMS_V) Then .ClearOP20_RMS_V = AsDoub(drw.ClearOP20_RMS_V)
                If Not Double.IsNaN(drw.ClearOP20_Volt_V) Then .ClearOP20_Volt_V = AsDoub(drw.ClearOP20_Volt_V)
                If Not Double.IsNaN(drw.ClearOP20_Freq_V) Then .ClearOP20_Freq_V = AsDoub(drw.ClearOP20_Freq_V)

                If Not Double.IsNaN(drw.ClearOP15_Power_V) Then .ClearOP15_Power_V = AsDoub(drw.ClearOP15_Power_V)
                If Not Double.IsNaN(drw.ClearOP15_RMS_V) Then .ClearOP15_RMS_V = AsDoub(drw.ClearOP15_RMS_V)
                If Not Double.IsNaN(drw.ClearOP15_Volt_V) Then .ClearOP15_Volt_V = AsDoub(drw.ClearOP15_Volt_V)
                If Not Double.IsNaN(drw.ClearOP15_Freq_V) Then .ClearOP15_Freq_V = AsDoub(drw.ClearOP15_Freq_V)

                If Not Double.IsNaN(drw.Future_1) Then .Future_1 = AsDoub(drw.Future_1)
                If Not Double.IsNaN(drw.Future_2) Then .Future_2 = AsDoub(drw.Future_2)
                If Not Double.IsNaN(drw.Future_3) Then .Future_3 = AsDoub(drw.Future_3)

                .Measurement_Grade = drw.Measurement_Grade
            End With

            Try
                DB_Con.SubmitChanges()
                retval = 1
            Catch
                Console.WriteLine("Update Didn't Work!")
                retval = -1
            End Try


        Else
            retval = -1
        End If

        Return retval

    End Function


    Private Function AsDoub(ByRef source As Double) As Double

        If Double.IsNaN(source) Then
            Return Nothing
        Else
            Return source
        End If

    End Function

    Private Sub PrintError(ByVal RecordID As String, ByVal field As String)

        Console.WriteLine(String.Format("Store to database error: Record = {0}, Field = {1}", RecordID, field))

    End Sub

End Module

