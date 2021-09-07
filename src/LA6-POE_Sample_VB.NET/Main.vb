Imports System.Net
Imports System.Net.Sockets
Imports System.Runtime.InteropServices

Module Main
    Private sock As Socket = Nothing

    ' product category
    Public Const PNS_PRODUCT_ID As UShort = &H4142

    ' PNS command identifier
    Private Const PNS_SMART_MODE_COMMAND As Byte = &H54         ' smart mode control command
    Private Const PNS_MUTE_COMMAND As Byte = &H4D               ' mute command
    Private Const PNS_STOP_PULSE_INPUT_COMMAND As Byte = &H50   ' stop/pulse input command
    Private Const PNS_RUN_CONTROL_COMMAND As Byte = &H53        ' operation control command
    Private Const PNS_DETAIL_RUN_CONTROL_COMMAND As Byte = &H44 ' detailed operation control command
    Private Const PNS_CLEAR_COMMAND As Byte = &H43              ' clear command
    Private Const PNS_REBOOT_COMMAND As Byte = &H42             ' reboot command
    Private Const PNS_GET_DATA_COMMAND As Byte = &H47           ' get status command
    Private Const PNS_GET_DETAIL_DATA_COMMAND As Byte = &H45    ' get detail status command

    ' response data for PNS commands
    Private Const PNS_ACK As Byte = &H6      ' normal response
    Private Const PNS_NAK As Byte = &H15     ' abnormal response

    ' mode
    Private Const PNS_LED_MODE As Byte = &H0    ' signal light mode
    Private Const PNS_SMART_MODE As Byte = &H1  ' smart mode

    ' LED unit for motion control command
    Private Const PNS_RUN_CONTROL_LED_OFF As Byte = &H0         ' light off
    Private Const PNS_RUN_CONTROL_LED_ON As Byte = &H1          ' light on
    Private Const PNS_RUN_CONTROL_LED_BLINKING As Byte = &H2    ' flashing
    Private Const PNS_RUN_CONTROL_LED_NO_CHANGE As Byte = &H9   ' no change

    ' buzzer for motion control command
    Private Const PNS_RUN_CONTROL_BUZZER_STOP As Byte = &H0         ' stop
    Private Const PNS_RUN_CONTROL_BUZZER_PATTERN1 As Byte = &H1     ' pattern 1
    Private Const PNS_RUN_CONTROL_BUZZER_PATTERN2 As Byte = &H2     ' pattern 2
    Private Const PNS_RUN_CONTROL_BUZZER_TONE As Byte = &H3         ' buzzer tone for simultaneous buzzer input
    Private Const PNS_RUN_CONTROL_BUZZER_NO_CHANGE As Byte = &H9    ' no change

    ' LED unit for detailed operation control command
    Private Const PNS_DETAIL_RUN_CONTROL_LED_OFF As Byte = &H0              ' light off
    Private Const PNS_DETAIL_RUN_CONTROL_LED_COLOR_RED As Byte = &H1        ' red
    Private Const PNS_DETAIL_RUN_CONTROL_LED_COLOR_YELLOW As Byte = &H2     ' yellow
    Private Const PNS_DETAIL_RUN_CONTROL_LED_COLOR_LEMON As Byte = &H3      ' Graduates
    Private Const PNS_DETAIL_RUN_CONTROL_LED_COLOR_GREEN As Byte = &H4      ' green
    Private Const PNS_DETAIL_RUN_CONTROL_LED_COLOR_SKY_BLUE As Byte = &H5   ' sky blue
    Private Const PNS_DETAIL_RUN_CONTROL_LED_COLOR_BLUE As Byte = &H6       ' blue
    Private Const PNS_DETAIL_RUN_CONTROL_LED_COLOR_PURPLE As Byte = &H7     ' purple
    Private Const PNS_DETAIL_RUN_CONTROL_LED_COLOR_PEACH As Byte = &H8      ' peach
    Private Const PNS_DETAIL_RUN_CONTROL_LED_COLOR_WHITE As Byte = &H9      ' white

    ' blinking action for detailed action control command
    Private Const PNS_DETAIL_RUN_CONTROL_BLINKING_OFF As Byte = &H0     ' blinking off
    Private Const PNS_DETAIL_RUN_CONTROL_BLINKING_ON As Byte = &H1      ' blinking ON

    ' buzzer for detailed action control command
    Private Const PNS_DETAIL_RUN_CONTROL_BUZZER_STOP As Byte = &H0          ' stop
    Private Const PNS_DETAIL_RUN_CONTROL_BUZZER_PATTERN1 As Byte = &H1      ' pattern 1
    Private Const PNS_DETAIL_RUN_CONTROL_BUZZER_PATTERN2 As Byte = &H2      ' pattern 2
    Private Const PNS_DETAIL_RUN_CONTROL_BUZZER_PATTERN3 As Byte = &H3      ' pattern 3
    Private Const PNS_DETAIL_RUN_CONTROL_BUZZER_PATTERN4 As Byte = &H4      ' pattern 4
    Private Const PNS_DETAIL_RUN_CONTROL_BUZZER_PATTERN5 As Byte = &H5      ' pattern 5
    Private Const PNS_DETAIL_RUN_CONTROL_BUZZER_PATTERN6 As Byte = &H6      ' pattern 6
    Private Const PNS_DETAIL_RUN_CONTROL_BUZZER_PATTERN7 As Byte = &H7      ' pattern 7
    Private Const PNS_DETAIL_RUN_CONTROL_BUZZER_PATTERN8 As Byte = &H8      ' pattern 8
    Private Const PNS_DETAIL_RUN_CONTROL_BUZZER_PATTERN9 As Byte = &H9      ' pattern 9
    Private Const PNS_DETAIL_RUN_CONTROL_BUZZER_PATTERN10 As Byte = &HA     ' pattern 10
    Private Const PNS_DETAIL_RUN_CONTROL_BUZZER_PATTERN11 As Byte = &HB     ' pattern 11

    ' operation control data structure
    Public Class PNS_RUN_CONTROL_DATA
        ' 1st LED unit pattern
        Public led1Pattern As Byte = 0

        ' 2nd LED unit pattern
        Public led2Pattern As Byte = 0

        ' 3rd LED unit pattern
        Public led3Pattern As Byte = 0

        ' 4th LED unit pattern
        Public led4Pattern As Byte = 0

        ' 5th LED unit pattern
        Public led5Pattern As Byte = 0

        ' buzzer pattern 1 to 3
        Public buzzerPattern As Byte = 0
    End Class

    ' detail operation control data structure
    Public Class PNS_DETAIL_RUN_CONTROL_DATA
        ' 1st color of LED unit
        Public led1Color As Byte = 0

        ' 2nd color of LED unit
        Public led2Color As Byte = 0

        ' 3rd color of LED unit
        Public led3Color As Byte = 0

        ' 4th color of LED unit
        Public led4Color As Byte = 0

        ' 5th color of LED unit
        Public led5Color As Byte = 0

        ' blinking action
        Public blinkingControl As Byte = 0

        ' buzzer pattern 1 to 11
        Public buzzerPattern As Byte = 0
    End Class

    ' status data of operation control
    Public Class PNS_STATUS_DATA
        ' input 1 to 8
        Public input As Byte() = New Byte(7) {}

        ' mode
        Public mode As Byte = 0

        ' status data when running signal light mode
        Public ledModeData As PNS_LED_MODE_DATA = Nothing

        ' status data during smart mode execution
        Public smartModeData As PNS_SMART_MODE_DATA = Nothing
    End Class

    ' status data when running in signal light mode
    Public Class PNS_LED_MODE_DATA
        ' 1st LED unit pattern
        Public led1Pattern As Byte = 0

        ' LED2nd LED unit pattern
        Public led2Pattern As Byte = 0

        ' 3rd LED unit pattern
        Public led3Pattern As Byte = 0

        ' 4th LED unit pattern
        Public led4Pattern As Byte = 0

        ' 5th LED unit pattern
        Public led5Pattern As Byte = 0

        ' buzzer patterns 1 through 11
        Public buzzerPattern As Byte = 0
    End Class

    ' state data when running smart mode
    Public Class PNS_SMART_MODE_DATA
        ' group number
        Public groupNo As Byte = 0

        ' mute
        Public mute As Byte = 0

        ' STOP input
        Public stopInput As Byte = 0

        ' pattern number
        Public patternNo As Byte = 0
    End Class

    ' status data of detailed operation control
    Public Class PNS_DETAIL_STATUS_DATA
        ' MAC address
        Public macAddress As Byte() = New Byte(5) {}

        ' Input 1 to 8
        Public input As Byte() = New Byte(7) {}

        ' mode
        Public mode As Byte = 0

        ' detailed status data when running signal light mode
        Public ledModeDetalData As PNS_LED_MODE_DETAIL_DATA = Nothing

        ' detailed state data when running in smart mode
        Public smartModeDetalData As PNS_SMART_MODE_DETAIL_DATA = Nothing
    End Class

    ' detailed state data when running in signal light mode
    Public Class PNS_LED_MODE_DETAIL_DATA
        ' 1st stage of LED unit
        Public ledUnit1Data As PNS_LED_UNIT_DATA = Nothing

        ' 2nd stage of LED unit
        Public ledUnit2Data As PNS_LED_UNIT_DATA = Nothing

        ' 3rd stage of LED unit
        Public ledUnit3Data As PNS_LED_UNIT_DATA = Nothing

        ' 4th stage of LED unit
        Public ledUnit4Data As PNS_LED_UNIT_DATA = Nothing

        ' 5th stage of LED unit
        Public ledUnit5Data As PNS_LED_UNIT_DATA = Nothing

        ' buzzer pattern 1 to 11
        Public buzzerPattern As Byte = 0
    End Class

    ' LED unit data
    Public Class PNS_LED_UNIT_DATA
        ' status
        Public ledPattern As Byte = 0

        ' R
        Public red As Byte = 0

        ' G
        Public green As Byte = 0

        ' B
        Public blue As Byte = 0
    End Class

    ' detail state data for smart mode execution
    Public Class PNS_SMART_MODE_DETAIL_DATA
        ' smart mode states
        Public smartModeData As PNS_SMART_MODE_DETAIL_STATE_DATA = Nothing

        ' 1st stage of LED unit
        Public ledUnit1Data As PNS_LED_UNIT_DATA = Nothing

        ' 2nd stage of LED unit
        Public ledUnit2Data As PNS_LED_UNIT_DATA = Nothing

        ' 3rd stage of LED unit
        Public ledUnit3Data As PNS_LED_UNIT_DATA = Nothing

        ' 4th stage of LED unit
        Public ledUnit4Data As PNS_LED_UNIT_DATA = Nothing

        ' 5th stage of LED unit
        Public ledUnit5Data As PNS_LED_UNIT_DATA = Nothing

        ' buzzer pattern 1 to 11
        Public buzzerPattern As Byte = 0
    End Class

    ' smart mode status data
    Public Class PNS_SMART_MODE_DETAIL_STATE_DATA
        ' group number
        Public groupNo As Byte = 0

        ' mute
        Public mute As Byte = 0

        ' STOP input
        Public stopInput As Byte = 0

        ' pattern number
        Public patternNo As Byte = 0

        ' last pattern
        Public lastPattern As Byte = 0
    End Class

    ' PHN command identifier
    Private Const PHN_WRITE_COMMAND As Byte = &H57  ' write command
    Private Const PHN_READ_COMMAND As Byte = &H52   ' read command

    ' response data for PNS command
    Private PHN_ACK() As Byte = {&H41, &H43, &H4B}  ' normal response
    Private PHN_NAK() As Byte = {&H4E, &H41, &H4B}  ' abnormal response

    ' action data of PHN command
    Private Const PHN_LED_UNIT1_BLINKING As Byte = &H20     ' 1st LED unit blinking
    Private Const PHN_LED_UNIT2_BLINKING As Byte = &H40     ' 2nd LED unit blinking
    Private Const PHN_LED_UNIT3_BLINKING As Byte = &H80     ' 3rd LED unit blinking
    Private Const PHN_BUZZER_PATTERN1 As Byte = &H8         ' buzzer pattern 1
    Private Const PHN_BUZZER_PATTERN2 As Byte = &H10        ' buzzer pattern 2
    Private Const PHN_LED_UNIT1_LIGHTING As Byte = &H1      ' 1st LED unit lighting
    Private Const PHN_LED_UNIT2_LIGHTING As Byte = &H2      ' 2nd LED unit lighting
    Private Const PHN_LED_UNIT3_LIGHTING As Byte = &H4      ' 3rd LED unit lighting

    ' <summary>
    ' Main Function
    ' </summary>
    Sub Main()
        Dim ret As Integer

        ' Connect to LA-POE
        ret = SocketOpen("192.168.10.1", 10000)
        If ret = -1 Then
            Return
        End If

        ' Get the command identifier specified by the command line argument
        Dim commandId As String = ""
        Dim cmds As String() = System.Environment.GetCommandLineArgs()
        If cmds.Length > 1 Then
            commandId = cmds(1)
        End If

        Select Case commandId
            Case "T"
                ' smart mode control command
                If cmds.Length >= 3 Then
                    PNS_SmartModeCommand(cmds(2))
                End If

            Case "M"
                ' mute command
                If cmds.Length >= 3 Then
                    PNS_MuteCommand(cmds(2))
                End If

            Case "P"
                ' stop/pulse input command
                If cmds.Length >= 3 Then
                    PNS_StopPulseInputCommand(cmds(2))
                End If

            Case "S"
                ' operation control command
                If cmds.Length >= 8 Then
                    Dim runControlData As PNS_RUN_CONTROL_DATA = New PNS_RUN_CONTROL_DATA With {
                        .led1Pattern = cmds(2),
                        .led2Pattern = cmds(3),
                        .led3Pattern = cmds(4),
                        .led4Pattern = cmds(5),
                        .led5Pattern = cmds(6),
                        .buzzerPattern = cmds(7)
                    }
                    PNS_RunControlCommand(runControlData)
                End If

            Case "D"
                ' detailed operation control command
                If cmds.Length >= 9 Then
                    Dim detalRunControlData As PNS_DETAIL_RUN_CONTROL_DATA = New PNS_DETAIL_RUN_CONTROL_DATA With {
                    .led1Color = cmds(2),
                    .led2Color = cmds(3),
                    .led3Color = cmds(4),
                    .led4Color = cmds(5),
                    .led5Color = cmds(6),
                    .blinkingControl = cmds(7),
                    .buzzerPattern = cmds(8)
                }
                    PNS_DetailRunControlCommand(detalRunControlData)
                End If

            Case "C"
                ' clear command
                PNS_ClearCommand()

            Case "B"
                ' reboot command
                If cmds.Length >= 3 Then
                    PNS_RebootCommand(cmds(2))
                End If

            Case "G"
                ' get status command
                Dim statusData As PNS_STATUS_DATA = New PNS_STATUS_DATA
                ret = PNS_GetDataCommand(statusData)
                If ret = 0 Then
                    ' Display acquired data
                    Console.WriteLine("Response data for status acquisition command")
                    ' Input1
                    Console.WriteLine("Input1 : " + statusData.input(0).ToString())
                    ' Input2
                    Console.WriteLine("Input2 : " + statusData.input(1).ToString())
                    ' Input3
                    Console.WriteLine("Input3 : " + statusData.input(2).ToString())
                    ' Input4
                    Console.WriteLine("Input4 : " + statusData.input(3).ToString())
                    ' Input5
                    Console.WriteLine("Input5 : " + statusData.input(4).ToString())
                    ' Input6
                    Console.WriteLine("Input6 : " + statusData.input(5).ToString())
                    ' Input7
                    Console.WriteLine("Input7 : " + statusData.input(6).ToString())
                    ' Input8
                    Console.WriteLine("Input8 : " + statusData.input(7).ToString())
                    ' mode
                    If statusData.mode = PNS_LED_MODE Then
                        ' signal light mode
                        Console.WriteLine("signal light mode")
                        ' 1st LED unit pattern
                        Console.WriteLine("1st LED unit pattern : " + statusData.ledModeData.led1Pattern.ToString())
                        ' 2nd LED unit pattern
                        Console.WriteLine("2nd LED unit pattern : " + statusData.ledModeData.led2Pattern.ToString())
                        ' 3rd LED unit pattern
                        Console.WriteLine("3rd LED unit pattern : " + statusData.ledModeData.led3Pattern.ToString())
                        ' 4th LED unit pattern
                        Console.WriteLine("4th LED unit pattern : " + statusData.ledModeData.led4Pattern.ToString())
                        ' 5th LED unit pattern
                        Console.WriteLine("5th LED unit pattern : " + statusData.ledModeData.led5Pattern.ToString())
                        ' buzzer pattern
                        Console.WriteLine("buzzer pattern: " + statusData.ledModeData.buzzerPattern.ToString())
                    Else
                        ' smart mode
                        Console.WriteLine("smart mode")
                        ' group number
                        Console.WriteLine("group number : " + statusData.smartModeData.groupNo.ToString())
                        ' mute
                        Console.WriteLine("mute : " + statusData.smartModeData.mute.ToString())
                        ' STOP input
                        Console.WriteLine("STOP input : " + statusData.smartModeData.stopInput.ToString())
                        ' pattern number
                        Console.WriteLine("pattern number : " + statusData.smartModeData.patternNo.ToString())
                    End If
                End If

            Case "E"
                ' get detail status command
                Dim detailStatusData As PNS_DETAIL_STATUS_DATA = New PNS_DETAIL_STATUS_DATA
                ret = PNS_GetDetailDataCommand(detailStatusData)
                If ret = 0 Then
                    ' Display acquired data
                    Console.WriteLine("Response data for status acquisition command")
                    ' MAC address
                    Console.WriteLine("MAC address : " + Convert.ToString(detailStatusData.macAddress(0), 16) + "-" _
                                                       + Convert.ToString(detailStatusData.macAddress(1), 16) + "-" _
                                                       + Convert.ToString(detailStatusData.macAddress(2), 16) + "-" _
                                                       + Convert.ToString(detailStatusData.macAddress(3), 16) + "-" _
                                                       + Convert.ToString(detailStatusData.macAddress(4), 16) + "-" _
                                                       + Convert.ToString(detailStatusData.macAddress(5), 16))
                    ' Input1
                    Console.WriteLine("Input1 : " + detailStatusData.input(0).ToString())
                    ' Input2
                    Console.WriteLine("Input2 : " + detailStatusData.input(1).ToString())
                    ' Input3
                    Console.WriteLine("Input3 : " + detailStatusData.input(2).ToString())
                    ' Input4
                    Console.WriteLine("Input4 : " + detailStatusData.input(3).ToString())
                    ' Input5
                    Console.WriteLine("Input5 : " + detailStatusData.input(4).ToString())
                    ' Input6
                    Console.WriteLine("Input6 : " + detailStatusData.input(5).ToString())
                    ' Input7
                    Console.WriteLine("Input7 : " + detailStatusData.input(6).ToString())
                    ' Input8
                    Console.WriteLine("Input8 : " + detailStatusData.input(7).ToString())
                    ' mode
                    If detailStatusData.mode = PNS_LED_MODE Then
                        ' signal light mode
                        Console.WriteLine("signal light mode")
                        ' 1st LED unit
                        Console.WriteLine("1st LED unit")
                        ' pattern
                        Console.WriteLine("pattern : " + detailStatusData.ledModeDetalData.ledUnit1Data.ledPattern.ToString())
                        ' R
                        Console.WriteLine("R : " + detailStatusData.ledModeDetalData.ledUnit1Data.red.ToString())
                        ' G
                        Console.WriteLine("G : " + detailStatusData.ledModeDetalData.ledUnit1Data.green.ToString())
                        ' B
                        Console.WriteLine("B : " + detailStatusData.ledModeDetalData.ledUnit1Data.blue.ToString())
                        ' 2nd LED unit
                        Console.WriteLine("2nd LED unit")
                        ' pattern
                        Console.WriteLine("pattern : " + detailStatusData.ledModeDetalData.ledUnit2Data.ledPattern.ToString())
                        ' R
                        Console.WriteLine("R : " + detailStatusData.ledModeDetalData.ledUnit2Data.red.ToString())
                        ' G
                        Console.WriteLine("G : " + detailStatusData.ledModeDetalData.ledUnit2Data.green.ToString())
                        ' B
                        Console.WriteLine("B : " + detailStatusData.ledModeDetalData.ledUnit2Data.blue.ToString())
                        ' 3rd LED unit
                        Console.WriteLine("3rd LED unit")
                        ' pattern
                        Console.WriteLine("pattern : " + detailStatusData.ledModeDetalData.ledUnit3Data.ledPattern.ToString())
                        ' R
                        Console.WriteLine("R : " + detailStatusData.ledModeDetalData.ledUnit3Data.red.ToString())
                        ' G
                        Console.WriteLine("G : " + detailStatusData.ledModeDetalData.ledUnit3Data.green.ToString())
                        ' B
                        Console.WriteLine("B : " + detailStatusData.ledModeDetalData.ledUnit3Data.blue.ToString())
                        ' 4th LED unit
                        Console.WriteLine("4th LED unit")
                        ' pattern
                        Console.WriteLine("pattern : " + detailStatusData.ledModeDetalData.ledUnit4Data.ledPattern.ToString())
                        ' R
                        Console.WriteLine("R : " + detailStatusData.ledModeDetalData.ledUnit4Data.red.ToString())
                        ' G
                        Console.WriteLine("G : " + detailStatusData.ledModeDetalData.ledUnit4Data.green.ToString())
                        ' B
                        Console.WriteLine("B : " + detailStatusData.ledModeDetalData.ledUnit4Data.blue.ToString())
                        ' 5th LED unit
                        Console.WriteLine("5th LED unit")
                        ' pattern
                        Console.WriteLine("pattern : " + detailStatusData.ledModeDetalData.ledUnit5Data.ledPattern.ToString())
                        ' R
                        Console.WriteLine("R : " + detailStatusData.ledModeDetalData.ledUnit5Data.red.ToString())
                        ' G
                        Console.WriteLine("G : " + detailStatusData.ledModeDetalData.ledUnit5Data.green.ToString())
                        ' B
                        Console.WriteLine("B : " + detailStatusData.ledModeDetalData.ledUnit5Data.blue.ToString())
                        ' buzzer pattern
                        Console.WriteLine("buzzer pattern: " + detailStatusData.ledModeDetalData.buzzerPattern.ToString())
                    Else
                        ' smart mode
                        Console.WriteLine("smart mode")
                        ' group number
                        Console.WriteLine("group number : " + detailStatusData.smartModeDetalData.smartModeData.groupNo.ToString())
                        ' mute
                        Console.WriteLine("mute : " + detailStatusData.smartModeDetalData.smartModeData.mute.ToString())
                        ' STOP input
                        Console.WriteLine("STOP input : " + detailStatusData.smartModeDetalData.smartModeData.stopInput.ToString())
                        ' pattern number
                        Console.WriteLine("pattern number : " + detailStatusData.smartModeDetalData.smartModeData.patternNo.ToString())
                        ' last pattern
                        Console.WriteLine("last pattern : " + detailStatusData.smartModeDetalData.smartModeData.lastPattern.ToString())
                        ' 1st LED unit
                        Console.WriteLine("1st LED unit")
                        ' pattern
                        Console.WriteLine("pattern : " + detailStatusData.smartModeDetalData.ledUnit1Data.ledPattern.ToString())
                        ' R
                        Console.WriteLine("R : " + detailStatusData.smartModeDetalData.ledUnit1Data.red.ToString())
                        ' G
                        Console.WriteLine("G : " + detailStatusData.smartModeDetalData.ledUnit1Data.green.ToString())
                        ' B
                        Console.WriteLine("B : " + detailStatusData.smartModeDetalData.ledUnit1Data.blue.ToString())
                        ' 2nd LED unit
                        Console.WriteLine("2nd LED unit")
                        ' pattern
                        Console.WriteLine("pattern : " + detailStatusData.smartModeDetalData.ledUnit2Data.ledPattern.ToString())
                        ' R
                        Console.WriteLine("R : " + detailStatusData.smartModeDetalData.ledUnit2Data.red.ToString())
                        ' G
                        Console.WriteLine("G : " + detailStatusData.smartModeDetalData.ledUnit2Data.green.ToString())
                        ' B
                        Console.WriteLine("B : " + detailStatusData.smartModeDetalData.ledUnit2Data.blue.ToString())
                        ' 3rd LED unit
                        Console.WriteLine("3rd LED unit")
                        ' pattern
                        Console.WriteLine("pattern : " + detailStatusData.smartModeDetalData.ledUnit3Data.ledPattern.ToString())
                        ' R
                        Console.WriteLine("R : " + detailStatusData.smartModeDetalData.ledUnit3Data.red.ToString())
                        ' G
                        Console.WriteLine("G : " + detailStatusData.smartModeDetalData.ledUnit3Data.green.ToString())
                        ' B
                        Console.WriteLine("B : " + detailStatusData.smartModeDetalData.ledUnit3Data.blue.ToString())
                        ' 4th LED unit
                        Console.WriteLine("4th LED unit")
                        ' pattern
                        Console.WriteLine("pattern : " + detailStatusData.smartModeDetalData.ledUnit4Data.ledPattern.ToString())
                        ' R
                        Console.WriteLine("R : " + detailStatusData.smartModeDetalData.ledUnit4Data.red.ToString())
                        ' G
                        Console.WriteLine("G : " + detailStatusData.smartModeDetalData.ledUnit4Data.green.ToString())
                        ' B
                        Console.WriteLine("B : " + detailStatusData.smartModeDetalData.ledUnit4Data.blue.ToString())
                        ' 5th LED unit
                        Console.WriteLine("5th LED unit")
                        ' pattern
                        Console.WriteLine("pattern : " + detailStatusData.smartModeDetalData.ledUnit5Data.ledPattern.ToString())
                        ' R
                        Console.WriteLine("R : " + detailStatusData.smartModeDetalData.ledUnit5Data.red.ToString())
                        ' G
                        Console.WriteLine("G : " + detailStatusData.smartModeDetalData.ledUnit5Data.green.ToString())
                        ' B
                        Console.WriteLine("B : " + detailStatusData.smartModeDetalData.ledUnit5Data.blue.ToString())
                        ' buzzer pattern
                        Console.WriteLine("buzzer pattern: " + detailStatusData.smartModeDetalData.buzzerPattern.ToString())
                    End If
                End If

            Case "W"
                ' write command
                If cmds.Length >= 3 Then
                    PHN_WriteCommand(cmds(2))
                End If

            Case "R"
                ' read command
                Dim runData As Byte
                ret = PHN_ReadCommand(runData)
                If ret = 0 Then
                    ' Display acquired data
                    Console.WriteLine("Response data for read command")
                    ' LED unit flashing
                    Console.WriteLine("LED unit flashing")
                    ' 1st LED unit
                    Console.WriteLine("1st LED unit : " + If(runData And &H20, 1, 0).ToString())
                    ' 2nd LED unit
                    Console.WriteLine("2nd LED unit : " + If(runData And &H40, 1, 0).ToString())
                    ' 3rd LED unit
                    Console.WriteLine("3rd LED unit : " + If(runData And &H80, 1, 0).ToString())
                    ' buzzer pattern
                    Console.WriteLine("buzzer pattern")
                    ' pattern1
                    Console.WriteLine("pattern1 : " + If(runData And &H8, 1, 0).ToString())
                    ' pattern2
                    Console.WriteLine("pattern2 : " + If(runData And &H10, 1, 0).ToString())
                    ' LED unit lighting
                    Console.WriteLine("LED unit lighting")
                    ' 1st LED unit
                    Console.WriteLine("1st LED unit : " + If(runData And &H1, 1, 0).ToString())
                    ' 2nd LED unit
                    Console.WriteLine("2nd LED unit : " + If(runData And &H2, 1, 0).ToString())
                    ' 3rd LED unit
                    Console.WriteLine("3rd LED unit : " + If(runData And &H4, 1, 0).ToString())
                End If
        End Select

        ' Close the socket
        SocketClose()
    End Sub

    ''' <summary>
    ''' Connect to LA-POE
    ''' </summary>
    ''' <param name="ip">IP address</param>
    ''' <param name="port">port number</param>
    ''' <returns>success: 0, failure: non-zero</returns>
    Public Function SocketOpen(ByVal ip As String, ByVal port As Integer) As Integer
        Try
            ' Set the IP address and port
            Dim ipAddress As IPAddress = IPAddress.Parse(ip)
            Dim remoteEP As IPEndPoint = New IPEndPoint(ipAddress, port)
            sock = New Socket(ipAddress.AddressFamily, SocketType.Stream, ProtocolType.Tcp)

            ' Create a socket
            If sock Is Nothing Then
                Console.WriteLine("failed to create socket")
                Return -1
            End If

            ' Connect to LA-POE
            sock.Connect(remoteEP)
        Catch ex As Exception
            Console.WriteLine(ex.Message)
            SocketClose()
            Return -1
        End Try

        Return 0
    End Function

    ''' <summary>
    ''' Close the socket.
    ''' </summary>
    Public Sub SocketClose()
        If sock IsNot Nothing Then
            ' Close the socket.
            sock.Shutdown(SocketShutdown.Both)
            sock.Close()
        End If
    End Sub

    ''' <summary>
    ''' Send command
    ''' </summary>
    ''' <param name="sendData">send data</param>
    ''' <param name="recvData">received data</param>
    ''' <returns>success: 0, failure: non-zero</returns>
    Public Function SendCommand(ByVal sendData As Byte(), ByRef recvData As Byte()) As Integer
        Dim ret As Integer
        recvData = Nothing

        Try

            If sock Is Nothing Then
                Console.WriteLine("socket is not")
                Return -1
            End If

            ' Send
            ret = sock.Send(sendData)
            If ret < 0 Then
                Console.WriteLine("failed to send")
                Return -1
            End If

            ' Receive response data
            Dim bytes As Byte() = New Byte(1023) {}
            Dim recvSize As Integer = sock.Receive(bytes)
            If recvSize < 0 Then
                Console.WriteLine("failed to recv")
                Return -1
            End If

            recvData = New Byte(recvSize - 1) {}
            Array.Copy(bytes, recvData, recvSize)
        Catch ex As Exception
            Console.WriteLine(ex.Message)
            Return -1
        End Try

        Return 0
    End Function

    ''' <summary>
    ''' Send smart mode control command for PNS command
    ''' Smart mode can be executed for the number specified in the data area
    ''' </summary>
    ''' <param name="groupNo">Group number to execute smart mode (0x01(Group No.1) to 0x1F(Group No.31))</param>
    ''' <returns>Success: 0, Failure: non-zero</returns>
    Public Function PNS_SmartModeCommand(ByVal groupNo As Byte) As Integer
        Dim ret As Integer

        Try
            Dim sendData As Byte() = {}

            ' Product Category (AB)
            sendData = sendData.Concat(BitConverter.GetBytes(PNS_PRODUCT_ID).Reverse()).ToArray()

            ' Command identifier (T)
            sendData = sendData.Concat(New Byte() {PNS_SMART_MODE_COMMAND}).ToArray()

            ' Empty (0)
            sendData = sendData.Concat(New Byte() {0}).ToArray()

            ' Data size
            sendData = sendData.Concat(BitConverter.GetBytes(CUShort(Marshal.SizeOf(groupNo))).Reverse()).ToArray()

            ' Data area
            sendData = sendData.Concat(New Byte() {groupNo}).ToArray()

            ' Send PNS commandz
            Dim recvData As Byte() = {}
            ret = SendCommand(sendData, recvData)
            If ret <> 0 Then
                Console.WriteLine("failed to send data")
                Return -1
            End If

            ' check the response data
            If recvData(0) = PNS_NAK Then
                ' receive abnormal response
                Console.WriteLine("negative acknowledge")
                Return -1
            End If

        Catch ex As Exception
            Console.WriteLine(ex.Message)
            Return -1
        End Try

        Return 0
    End Function

    ''' <summary>
    ''' Send mute command for PNS command
    ''' Can control the buzzer ON/OFF while Smart Mode is running
    ''' </summary>
    ''' <param name="mute">Buzzer ON/OFF (ON: 1, OFF: 0)</param>
    ''' <returns>success: 0, failure: non-zero</returns>
    Public Function PNS_MuteCommand(ByVal mute As Byte) As Integer
        Dim ret As Integer

        Try
            Dim sendData As Byte() = {}

            ' Product Category (AB)
            sendData = sendData.Concat(BitConverter.GetBytes(PNS_PRODUCT_ID).Reverse()).ToArray()

            ' Command identifier (M)
            sendData = sendData.Concat(New Byte() {PNS_MUTE_COMMAND}).ToArray()

            ' Empty (0)
            sendData = sendData.Concat(New Byte() {0}).ToArray()

            ' Data size
            sendData = sendData.Concat(BitConverter.GetBytes(CUShort(Marshal.SizeOf(mute))).Reverse()).ToArray()

            ' Data area
            sendData = sendData.Concat(New Byte() {mute}).ToArray()

            ' Send PNS command
            Dim recvData As Byte() = {}
            ret = SendCommand(sendData, recvData)
            If ret <> 0 Then
                Console.WriteLine("failed to send data")
                Return -1
            End If

            ' check the response data
            If recvData(0) = PNS_NAK Then
                ' receive abnormal response
                Console.WriteLine("negative acknowledge")
                Return -1
            End If

        Catch ex As Exception
            Console.WriteLine(ex.Message)
            Return -1
        End Try

        Return 0
    End Function

    ''' <summary>
    ''' Send stop/pulse input command for PNS command
    ''' Transmit during time trigger mode operation to control stop/resume of pattern (STOP input)
    ''' Sending this command during pulse trigger mode operation enables pattern transition (trigger input).
    ''' </summary>
    ''' <param name="input">STOP input/trigger input (STOP input ON/trigger input: 1, STOP input: 0)</param>
    ''' <returns>Success: 0, failure: non-zero</returns>
    Public Function PNS_StopPulseInputCommand(ByVal input As Byte) As Integer
        Dim ret As Integer

        Try
            Dim sendData As Byte() = {}

            ' Product Category (AB)
            sendData = sendData.Concat(BitConverter.GetBytes(PNS_PRODUCT_ID).Reverse()).ToArray()

            ' Command identifier (P)
            sendData = sendData.Concat(New Byte() {PNS_STOP_PULSE_INPUT_COMMAND}).ToArray()

            ' Empty (0)
            sendData = sendData.Concat(New Byte() {0}).ToArray()

            ' Data size
            sendData = sendData.Concat(BitConverter.GetBytes(CUShort(Marshal.SizeOf(input))).Reverse()).ToArray()

            ' Data area
            sendData = sendData.Concat(New Byte() {input}).ToArray()

            ' Send PNS command
            Dim recvData As Byte() = {}
            ret = SendCommand(sendData, recvData)
            If ret <> 0 Then
                Console.WriteLine("failed to send data")
                Return -1
            End If

            ' check the response data
            If recvData(0) = PNS_NAK Then
                ' receive abnormal response
                Console.WriteLine("negative acknowledge")
                Return -1
            End If

        Catch ex As Exception
            Console.WriteLine(ex.Message)
            Return -1
        End Try

        Return 0
    End Function

    ''' <summary>
    ''' Send operation control command for PNS command
    ''' Each stage of the LED unit and the buzzer (1 to 3) can be controlled by the pattern specified in the data area
    ''' Operates with the color and buzzer set in the signal light mode
    ''' </summary>
    ''' <param name="runControlData">
    ''' Pattern of the 1st to 5th stage of the LED unit and buzzer (1 to 3)
    ''' Pattern of LED unit (off: 0, on: 1, blinking: 2, no change: 9)
    ''' Pattern of buzzer (stop: 0, pattern 1: 1, pattern 2: 2, buzzer tone when input simultaneously with buzzer: 3, no change: 9)
    ''' </param>
    ''' <returns>success: 0, failure: non-zero</returns>
    Public Function PNS_RunControlCommand(ByVal runControlData As PNS_RUN_CONTROL_DATA) As Integer
        Dim ret As Integer

        Try
            Dim sendData As Byte() = {}

            ' Product Category (AB)
            sendData = sendData.Concat(BitConverter.GetBytes(PNS_PRODUCT_ID).Reverse()).ToArray()

            ' Command identifier (S)
            sendData = sendData.Concat(New Byte() {PNS_RUN_CONTROL_COMMAND}).ToArray()

            ' Empty (0)
            sendData = sendData.Concat(New Byte() {0}).ToArray()

            ' data size, data area
            Dim data As Byte() = {
                runControlData.led1Pattern,     ' 1st LED unit pattern
                runControlData.led2Pattern,     ' 2nd LED unit pattern
                runControlData.led3Pattern,     ' 3rd LED unit pattern
                runControlData.led4Pattern,     ' 4th LED unit pattern
                runControlData.led5Pattern,     ' 5th LED unit pattern
                runControlData.buzzerPattern    ' Buzzer pattern 1 to 3
            }
            sendData = sendData.Concat(BitConverter.GetBytes(CUShort(data.Length)).Reverse()).ToArray()
            sendData = sendData.Concat(data).ToArray()

            ' Send PNS command
            Dim recvData As Byte() = {}
            ret = SendCommand(sendData, recvData)
            If ret <> 0 Then
                Console.WriteLine("failed to send data")
                Return -1
            End If

            ' check the response data
            If recvData(0) = PNS_NAK Then
                ' receive abnormal response
                Console.WriteLine("negative acknowledge")
                Return -1
            End If

        Catch ex As Exception
            Console.WriteLine(ex.Message)
            Return -1
        End Try

        Return 0
    End Function

    ''' <summary>
    ''' Send detailed operation control command for PNS command
    ''' The color and operation pattern of each stage of the LED unit and the buzzer pattern (1 to 11) can be specified and controlled in the data area
    ''' </summary>
    ''' <param name="detailRunControlData">
    ''' Pattern of the 1st to 5th stage of the LED unit, blinking operation and buzzer (1 to 3)
    ''' Pattern of LED unit (off: 0, red: 1, yellow: 2, lemon: 3, green: 4, sky blue: 5, blue: 6, purple: 7, peach: 8, white: 9)
    ''' Flashing action (Flashing OFF: 0, Flashing ON: 1)
    ''' Buzzer pattern (Stop: 0, Pattern 1: 1, Pattern 2: 2, Pattern 3: 3, Pattern 4: 4, Pattern 5: 5, Pattern 6: 6, Pattern 7: 7, Pattern 8: 8, Pattern 9: 9, Pattern 10: 10, Pattern 11: 11)
    ''' </param>
    ''' <returns>success: 0, failure: non-zero</returns>
    Public Function PNS_DetailRunControlCommand(ByVal detailRunControlData As PNS_DETAIL_RUN_CONTROL_DATA) As Integer
        Dim ret As Integer

        Try
            Dim sendData As Byte() = {}

            ' Product Category (AB)
            sendData = sendData.Concat(BitConverter.GetBytes(PNS_PRODUCT_ID).Reverse()).ToArray()

            ' Command identifier (D)
            sendData = sendData.Concat(New Byte() {PNS_DETAIL_RUN_CONTROL_COMMAND}).ToArray()

            ' Empty (0)
            sendData = sendData.Concat(New Byte() {0}).ToArray()

            ' data size, data area
            Dim data As Byte() = {
                detailRunControlData.led1Color,         ' 1st color of LED unit
                detailRunControlData.led2Color,         ' 2nd color of LED unit
                detailRunControlData.led3Color,         ' 3rd color of LED unit
                detailRunControlData.led4Color,         ' 4th color of LED unit
                detailRunControlData.led5Color,         ' 5th color of LED unit
                detailRunControlData.blinkingControl,   ' blinking operation
                detailRunControlData.buzzerPattern      ' buzzer pattern 1 to 11
            }
            sendData = sendData.Concat(BitConverter.GetBytes(CUShort(data.Length)).Reverse()).ToArray()
            sendData = sendData.Concat(data).ToArray()

            ' Send PNS command
            Dim recvData As Byte() = {}
            ret = SendCommand(sendData, recvData)
            If ret <> 0 Then
                Console.WriteLine("failed to send data")
                Return -1
            End If

            ' check the response data
            If recvData(0) = PNS_NAK Then
                ' receive abnormal response
                Console.WriteLine("negative acknowledge")
                Return -1
            End If

        Catch ex As Exception
            Console.WriteLine(ex.Message)
            Return -1
        End Try

        Return 0
    End Function

    ''' <summary>
    ''' Send clear command for PNS command
    ''' Turn off the LED unit and stop the buzzer
    ''' </summary>
    ''' <returns>success: 0, failure: non-zero</returns>
    Public Function PNS_ClearCommand() As Integer
        Dim ret As Integer

        Try
            Dim sendData As Byte() = {}

            ' Product Category (AB)
            sendData = sendData.Concat(BitConverter.GetBytes(PNS_PRODUCT_ID).Reverse()).ToArray()

            ' Command identifier (C)
            sendData = sendData.Concat(New Byte() {PNS_CLEAR_COMMAND}).ToArray()

            ' Empty (0)
            sendData = sendData.Concat(New Byte() {0}).ToArray()

            ' Data size
            sendData = sendData.Concat(BitConverter.GetBytes(CUShort(0))).ToArray()

            ' Send PNS command
            Dim recvData As Byte() = {}
            ret = SendCommand(sendData, recvData)
            If ret <> 0 Then
                Console.WriteLine("failed to send data")
                Return -1
            End If

            ' check the response data
            If recvData(0) = PNS_NAK Then
                ' receive abnormal response
                Console.WriteLine("negative acknowledge")
                Return -1
            End If

        Catch ex As Exception
            Console.WriteLine(ex.Message)
            Return -1
        End Try

        Return 0
    End Function

    ''' <summary>
    ''' Send restart command for PNS command
    ''' LA6-POE can be restarted
    ''' </summary>
    ''' <param name="password">Password set in the password setting of Web Configuration</param>.
    ''' <returns>success: 0, failure: non-zero</returns>
    Public Function PNS_RebootCommand(ByVal password As String) As Integer
        Dim ret As Integer

        Try
            Dim sendData As Byte() = {}

            ' Product Category (AB)
            sendData = sendData.Concat(BitConverter.GetBytes(PNS_PRODUCT_ID).Reverse()).ToArray()

            ' Command identifier (B)
            sendData = sendData.Concat(New Byte() {PNS_REBOOT_COMMAND}).ToArray()

            ' Empty (0)
            sendData = sendData.Concat(New Byte() {0}).ToArray()

            ' Data size
            sendData = sendData.Concat(BitConverter.GetBytes(CUShort(password.Length)).Reverse()).ToArray()

            ' Data area
            sendData = sendData.Concat(System.Text.Encoding.ASCII.GetBytes(password)).ToArray()

            ' Send PNS command
            Dim recvData As Byte() = {}
            ret = SendCommand(sendData, recvData)
            If ret <> 0 Then
                Console.WriteLine("failed to send data")
                Return -1
            End If

            ' check the response data
            If recvData(0) = PNS_NAK Then
                ' receive abnormal response
                Console.WriteLine("negative acknowledge")
                Return -1
            End If

        Catch ex As Exception
            Console.WriteLine(ex.Message)
            Return -1
        End Try

        Return 0
    End Function

    ''' <summary>
    ''' Send status acquisition command for PNS command
    ''' Signal line/contact input status and LED unit and buzzer status can be acquired
    ''' </summary>
    ''' <param name="statusData">Received data of status acquisition command (status of signal line/contact input and status of LED unit and buzzer)</param>
    ''' <returns>Success: 0, failure: non-zero</returns>
    Public Function PNS_GetDataCommand(ByRef statusData As PNS_STATUS_DATA) As Integer
        Dim ret As Integer
        statusData = New PNS_STATUS_DATA()

        Try
            Dim sendData As Byte() = {}

            ' Product Category (AB)
            sendData = sendData.Concat(BitConverter.GetBytes(PNS_PRODUCT_ID).Reverse()).ToArray()

            ' Command identifier (G)
            sendData = sendData.Concat(New Byte() {PNS_GET_DATA_COMMAND}).ToArray()

            ' Empty (0)
            sendData = sendData.Concat(New Byte() {0}).ToArray()

            ' Data size
            sendData = sendData.Concat(BitConverter.GetBytes(CShort(0)).Reverse()).ToArray()

            ' Send PNS command
            Dim recvData As Byte() = {}
            ret = SendCommand(sendData, recvData)
            If ret <> 0 Then
                Console.WriteLine("failed to send data")
                Return -1
            End If

            ' check the response data
            If recvData(0) = PNS_NAK Then
                ' receive abnormal response
                Console.WriteLine("negative acknowledge")
                Return -1
            End If

            ' Input 1 to 8
            statusData.input = New Byte(7) {}
            Array.Copy(recvData, statusData.input, statusData.input.Length)

            ' Mode
            statusData.mode = recvData(8)

            ' Check the mode
            If statusData.mode = PNS_LED_MODE Then
                ' signal light mode
                statusData.ledModeData = New PNS_LED_MODE_DATA With {
                    .led1Pattern = recvData(9),     ' 1st LED unit pattern
                    .led2Pattern = recvData(10),    ' 2nd LED unit pattern
                    .led3Pattern = recvData(11),    ' 3rd LED unit pattern
                    .led4Pattern = recvData(12),    ' 4th LED unit pattern
                    .led5Pattern = recvData(13),    ' 5th LED unit pattern
                    .buzzerPattern = recvData(14)   ' buzzer pattern 1 to 11
                }
            Else
                ' smart mode
                statusData.smartModeData = New PNS_SMART_MODE_DATA With {
                    .groupNo = recvData(9),         ' group number
                    .mute = recvData(10),           ' mute
                    .stopInput = recvData(11),      ' STOP input
                    .patternNo = recvData(12)       ' pattern number
                }
            End If

        Catch ex As Exception
            Console.WriteLine(ex.Message)
            Return -1
        End Try

        Return 0
    End Function

    ''' <summary>
    ''' Send command to get detailed status of PNS command
    ''' Signal line/contact input status, LED unit and buzzer status, and color information for each stage can be acquired
    ''' </summary>
    ''' <param name="detailStatusData">Received data of detail status acquisition command (status of signal line/contact input, status of LED unit and buzzer, and color information of each stage)</param>
    ''' <returns>Success: 0, failure: non-zero</returns>
    Public Function PNS_GetDetailDataCommand(ByRef detailStatusData As PNS_DETAIL_STATUS_DATA) As Integer
        Dim ret As Integer
        detailStatusData = New PNS_DETAIL_STATUS_DATA()

        Try
            Dim sendData As Byte() = {}

            ' Product Category (AB)
            sendData = sendData.Concat(BitConverter.GetBytes(PNS_PRODUCT_ID).Reverse()).ToArray()

            ' Command identifier (E)
            sendData = sendData.Concat(New Byte() {PNS_GET_DETAIL_DATA_COMMAND}).ToArray()

            ' Empty (0)
            sendData = sendData.Concat(New Byte() {0}).ToArray()

            ' Data size
            sendData = sendData.Concat(BitConverter.GetBytes(CShort(0)).Reverse()).ToArray()

            ' Send PNS command
            Dim recvData As Byte() = {}
            ret = SendCommand(sendData, recvData)
            If ret <> 0 Then
                Console.WriteLine("failed to send data")
                Return -1
            End If

            ' check the response data
            If recvData(0) = PNS_NAK Then
                ' receive abnormal response
                Console.WriteLine("negative acknowledge")
                Return -1
            End If

            ' MAC Address
            detailStatusData.macAddress = New Byte(5) {}
            Array.Copy(recvData, detailStatusData.macAddress, detailStatusData.macAddress.Length)

            ' Input 1 to 8
            detailStatusData.input = New Byte(7) {}
            Array.Copy(recvData, 6, detailStatusData.input, 0, detailStatusData.input.Length)

            ' Mode
            detailStatusData.mode = recvData(14)

            ' Check the mode
            If detailStatusData.mode = PNS_LED_MODE Then
                ' signal light mode
                detailStatusData.ledModeDetalData = New PNS_LED_MODE_DETAIL_DATA With {
                    .ledUnit1Data = New PNS_LED_UNIT_DATA With {    ' 1st stage of LED unit
                        .ledPattern = recvData(19), ' state
                        .red = recvData(20),        ' R
                        .green = recvData(21),      ' G
                        .blue = recvData(22)        ' B
                    },
                    .ledUnit2Data = New PNS_LED_UNIT_DATA With {    ' 2nd stage of LED unit
                        .ledPattern = recvData(23), ' state
                        .red = recvData(24),        ' R
                        .green = recvData(25),      ' G
                        .blue = recvData(26)        ' B
                    },
                    .ledUnit3Data = New PNS_LED_UNIT_DATA With {    ' 3rd stage of LED unit
                        .ledPattern = recvData(27), ' state
                        .red = recvData(28),        ' R
                        .green = recvData(29),      ' G
                        .blue = recvData(30)        ' B
                    },
                    .ledUnit4Data = New PNS_LED_UNIT_DATA With {    ' 4th stage of LED unit
                        .ledPattern = recvData(31), ' state
                        .red = recvData(32),        ' R
                        .green = recvData(33),      ' G
                        .blue = recvData(34)        ' B
                    },
                    .ledUnit5Data = New PNS_LED_UNIT_DATA With {    ' 5th stage of LED unit
                        .ledPattern = recvData(35), ' state
                        .red = recvData(36),        ' R
                        .green = recvData(37),      ' G
                        .blue = recvData(38)        ' B
                    },
                    .buzzerPattern = recvData(39)   ' buzzer patterns 1-11
                }
            Else
                detailStatusData.smartModeDetalData = New PNS_SMART_MODE_DETAIL_DATA With {
                    .smartModeData = New PNS_SMART_MODE_DETAIL_STATE_DATA With {    ' smart mode status
                        .groupNo = recvData(19),    ' group number
                        .mute = recvData(20),       ' mute
                        .stopInput = recvData(21),  ' STOP input
                        .patternNo = recvData(22),  ' pattern number
                        .lastPattern = recvData(23) ' last pattern
                    },
                    .ledUnit1Data = New PNS_LED_UNIT_DATA With {    ' 1st stage of LED unit
                        .ledPattern = recvData(24), ' state
                        .red = recvData(25),        ' R
                        .green = recvData(26),      ' G
                        .blue = recvData(27)        ' B
                    },
                    .ledUnit2Data = New PNS_LED_UNIT_DATA With {    ' 2nd stage of LED unit
                        .ledPattern = recvData(28), ' state
                        .red = recvData(29),        ' R
                        .green = recvData(30),      ' G
                        .blue = recvData(31)        ' B
                    },
                    .ledUnit3Data = New PNS_LED_UNIT_DATA With {    ' 3rd stage of LED unit
                        .ledPattern = recvData(32), ' state
                        .red = recvData(33),        ' R
                        .green = recvData(34),      ' G
                        .blue = recvData(35)        ' B
                    },
                    .ledUnit4Data = New PNS_LED_UNIT_DATA With {    ' 4th stage of LED unit
                        .ledPattern = recvData(36), ' state
                        .red = recvData(37),        ' R
                        .green = recvData(38),      ' G
                        .blue = recvData(39)        ' B
                    },
                    .ledUnit5Data = New PNS_LED_UNIT_DATA With {    ' 5th stage of LED unit
                        .ledPattern = recvData(40), ' state
                        .red = recvData(41),        ' R
                        .green = recvData(42),      ' G
                        .blue = recvData(43)        ' B
                    },
                    .buzzerPattern = recvData(44)   ' buzzer patterns 1-11
                }
            End If

        Catch ex As Exception
            Console.WriteLine(ex.Message)
            Return -1
        End Try

        Return 0
    End Function

    ''' <summary>
    ''' Send PHN command write command
    ''' Can control the lighting and blinking of LED units 1 to 3 stages, and buzzer patterns 1 and 2
    ''' </summary>
    ''' <param name="runData">
    ''' Operation data for lighting and blinking of LED unit 1 to 3 stages, and buzzer pattern 1 and 2
    ''' bit7: 3rd LED unit blinking (OFF: 0, ON: 1)
    ''' bit6: 2nd LED unit blinking (OFF: 0, ON: 1)
    ''' bit5: 1st LED unit blinking (OFF: 0, ON: 1)
    ''' bit4: Buzzer pattern 2 (OFF: 0, ON: 1)
    ''' bit3: Buzzer pattern 1 (OFF: 0, ON: 1)
    ''' bit2: 3rd LED unit lighting (OFF: 0, ON: 1)
    ''' bit1: 2nd LED unit lighting (OFF: 0, ON: 1)
    ''' bit0: 1st LED unit lighting (OFF: 0, ON: 1)
    ''' </param>
    ''' <returns>success: 0, failure: non-zero</returns>
    Public Function PHN_WriteCommand(ByVal runData As Byte) As Integer
        Dim ret As Integer

        Try
            Dim sendData As Byte() = {}

            ' Command identifier (W)
            sendData = sendData.Concat(New Byte() {PHN_WRITE_COMMAND}).ToArray()

            ' Operation data
            sendData = sendData.Concat(New Byte() {runData}).ToArray()

            ' send PHN command
            Dim recvData As Byte() = {}
            ret = SendCommand(sendData, recvData)
            If ret <> 0 Then
                Console.WriteLine("failed to send data")
                Return -1
            End If

            ' check the response data
            If recvData.SequenceEqual(PHN_NAK) Then
                ' receive abnormal response
                Console.WriteLine("negative acknowledge")
                Return -1
            End If

        Catch ex As Exception
            Console.WriteLine(ex.Message)
            Return -1
        End Try

        Return 0
    End Function

    ''' <summary>
    ''' Send command to read PHN command
    ''' Get information about LED unit 1 to 3 stage lighting and blinking, and buzzer pattern 1 and 2
    ''' </summary>
    ''' <param name="runData">Received data of read command (operation data of LED unit 1 to 3 stages lighting and blinking, buzzer pattern 1,2)</param>
    ''' <returns>Success: 0, failure: non-zero</returns>
    Public Function PHN_ReadCommand(ByRef runData As Byte) As Integer
        Dim ret As Integer
        runData = 0

        Try
            Dim sendData As Byte() = {}

            ' Command identifier (R)
            sendData = sendData.Concat(New Byte() {PHN_READ_COMMAND}).ToArray()

            ' send PHN command
            Dim recvData As Byte() = {}
            ret = SendCommand(sendData, recvData)
            If ret <> 0 Then
                Console.WriteLine("failed to send data")
                Return -1
            End If

            ' check the response data
            If recvData(0) <> PHN_READ_COMMAND Then
                ' receive abnormal response
                Console.WriteLine("negative acknowledge")
                Return -1
            End If

            ' Response data
            runData = recvData(1)
        Catch ex As Exception
            Console.WriteLine(ex.Message)
            Return -1
        End Try

        Return 0
    End Function

End Module
