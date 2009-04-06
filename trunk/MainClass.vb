Imports System
Imports System.Windows
Imports System.Windows.Forms
Public Class csMovieLibrary
#Region "mciSendString"
    Private Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Integer, ByVal lpstrBuffer As String, ByVal uLength As Integer) As Integer 'Get the error message of the mcidevice if any

    Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Integer, ByVal hwndCallback As Integer) As Integer 'Send command strings to the mci device

    Private Declare Function getTickCount Lib "kernel32" Alias "GetTickCount" () As Integer

#End Region

#Region "Fields"
    Public retVal As Integer ' used to store our return value from the mci interface
    Private retData As String = Space$(128) ' used to store our return data from various commands

    Private _filename As String = Nothing ' used to store our file to play
    Private _checkError As String = Nothing
    Private _audioStatus As String = Nothing
    Private _getCurrentSize As String = Nothing
    Private _bitsPerPixel As Integer = Nothing
    Private _movieInputSource As String = Nothing
    Private _movieOutputSource As String = Nothing
    Private _isMoviePlaying As Boolean = Nothing
    Private _deviceName As String = Nothing
    Private _deviceVersion As String = Nothing
    Private _nominalFrameRate As Integer = Nothing
    Private _framePerSecondRate As Integer = Nothing
    Private _getDefaultSize As String = Nothing
    Private _durationInFrames As Integer = Nothing
    Private _durationInMS As Integer = Nothing
    Private _durationInSec As Integer = Nothing
    Private _positionInFrames As Integer = Nothing
    Private _positionInMS As Integer = Nothing
    Private _positionInSec As Integer = Nothing
    Private _getPlayingStatus As String = Nothing
    Private _getCurrentTimeFormat As String = Nothing
    Private _leftChannelVolume As Integer = Nothing
    Private _rightChannelVolume As Integer = Nothing
    Private _isDeviceReady As Boolean = Nothing
    Private _formatPosition As String = Nothing
    Private _formatDuration As String = Nothing
    Private _volumeLevel As Integer = Nothing
#End Region

#Region "Methods and properties"


    Public Property filename() As String
        Get
            Try
                filename = _filename
            Catch exc As Exception
                MessageBox.Show(exc.Message, " Error", MessageBoxButtons.OK)
            End Try
        End Get
        Set(ByVal Value As String)
            Try
                _filename = Chr(34) + Value + Chr(34)
            Catch exc As Exception
                MessageBox.Show(exc.Message, " Error", MessageBoxButtons.OK)
            End Try

        End Set

    End Property

    Public WriteOnly Property stepFrames() As Integer
        Set(ByVal Value As Integer)
            Try
                retVal = mciSendString("step movie by " & Value, 0, 0, 0)
            Catch exc As Exception
                MessageBox.Show(exc.Message, " Error", MessageBoxButtons.OK)
            End Try
        End Set
    End Property

    Public Sub restoreSizeDefault()
        Try

            retVal = mciSendString("put movie window", 0, 0, 0)

        Catch exc As Exception

            MessageBox.Show(exc.Message, " Error", MessageBoxButtons.OK)


        End Try
    End Sub

    Public Sub openMovie()
        Try
            retVal = mciSendString("close movie", 0, 0, 0)
            retVal = mciSendString("open " & filename & " type mpegvideo alias movie", 0, 0, 0)
        Catch exc As Exception
            MessageBox.Show(exc.Message, " Error", MessageBoxButtons.OK)
        End Try
    End Sub

    Public Sub openMovieInWindow(ByVal hWnd As Integer)
        Try
            retVal = mciSendString("close movie", 0, 0, 0)
            retVal = mciSendString("open " & filename & " type mpegvideo alias movie parent " & hWnd & " style " & "child" & " ", 0, 0, 0)
        Catch exc As Exception
            MessageBox.Show(exc.Message, " Error", MessageBoxButtons.OK)
        End Try
    End Sub

    Public Sub minimizeMovie()
        Try
            retVal = mciSendString("window movie state minimized", 0, 0, 0)
        Catch exc As Exception
            MessageBox.Show(exc.Message, " Error", MessageBoxButtons.OK)
        End Try
    End Sub

    Public Sub playMovie()
        'play the movie after you open it
        Try
            retVal = mciSendString("play movie", 0, 0, 0)
        Catch exc As Exception
            MessageBox.Show(exc.Message, " Error", MessageBoxButtons.OK)
        End Try
    End Sub

    Public WriteOnly Property hideMovie() As Boolean
        'true = hide the video
        'false = show the video
        Set(ByVal Value As Boolean)
            Try
                If Value = True Then
                    'hides the movie window
                    retVal = mciSendString("window movie state hide", 0, 0, 0)

                Else

                    'will show the window if it was hidden with the hideMovie function
                    retVal = mciSendString("window movie state show", 0, 0, 0)

                End If

            Catch exc As Exception

                MessageBox.Show(exc.Message, " Error", MessageBoxButtons.OK)

            End Try

        End Set

    End Property

    Public Function restoreOriginalMovieState()

        'will restore the window to its original state

        Try

            retVal = mciSendString("window movie state restore", 0, 0, 0)

        Catch exc As Exception

            MessageBox.Show(exc.Message, " Error", MessageBoxButtons.OK)

        End Try

    End Function

    Public Sub stopMovie()

        'stops the playing of the movie

        Try

            retVal = mciSendString("stop movie", 0, 0, 0)

        Catch exc As Exception

            MessageBox.Show(exc.Message, " Error", MessageBoxButtons.OK)

        End Try

    End Sub

    Public ReadOnly Property getCurrentSizeLocation() As String

        Get

            'returns the current width, height and location of the movie

            Try

                retVal = mciSendString("where movie destination max", retData, 128, 0)

                _getCurrentSize = retData

                getCurrentSizeLocation = _getCurrentSize

                _getCurrentSize = Nothing

            Catch exc As Exception

                MessageBox.Show(exc.Message, " Error", MessageBoxButtons.OK)

            End Try

        End Get

    End Property

    Public Sub extractCurrentSize(ByVal wWidth As Integer, ByVal wHeight As Integer)

        'returns the current size parameters of the movie

        Try

            Dim a As String = Nothing
            Dim b As String = Nothing
            Dim C As String = Nothing
            Dim f As String = Nothing
            Dim g As String = Nothing

            a = getCurrentSizeLocation

            b = InStr(1, a, " ")

            f = Mid(a, CDbl(C) + 1)

            g = CStr(InStr(1, f, " "))

            wWidth = Val(Left(f, CInt(g))) 'width

            wHeight = Val(Mid(f, CInt(g))) 'height

            a = Nothing
            b = Nothing
            C = Nothing
            f = Nothing
            g = Nothing

        Catch exc As Exception

            MessageBox.Show(exc.Message, " Error", MessageBoxButtons.OK)

        End Try

    End Sub

    Public Sub extractDefaultSize(ByVal wWidth As Integer, ByVal wHeight As Integer)

        'returns the default size of the movie. even if the size of the movie has 
        'been(changed)

        Try

            Dim a As String = Nothing
            Dim b As String = Nothing
            Dim c As String = Nothing
            Dim f As String = Nothing
            Dim g As String = Nothing

            a = getDefaultSizeLocation

            b = CStr(InStr(1, a, " ")) '2

            'C = CStr(InStr(CDbl(b) + 1, a, " ")) '4

            f = Mid(a, CDbl(c) + 1) '9

            g = CStr(InStr(1, f, " "))

            wWidth = Val(Left(f, CInt(g))) 'width
            wHeight = Val(Mid(f, CInt(g))) 'height

            a = Nothing
            b = Nothing
            c = Nothing
            f = Nothing
            g = Nothing

        Catch exc As Exception

            MessageBox.Show(exc.Message, " Error", MessageBoxButtons.OK)

        End Try

    End Sub

    Public ReadOnly Property bitsPerPixel() As Integer

        Get

            'will return the bitsperpixels of certain AVI movies only

            Try

                retVal = mciSendString("status movie bitsperpel", retData, 128, 0)

                _bitsPerPixel = Val(retData)

                bitsPerPixel = _bitsPerPixel

                _bitsPerPixel = Nothing

            Catch exc As Exception

                MessageBox.Show(exc.Message, " Error", MessageBoxButtons.OK)

            End Try

        End Get

    End Property

    Public ReadOnly Property movieInputSource() As String

        Get

            'returns the current input source ex: file

            Try

                retVal = mciSendString("status movie monitor input", retData, 128, 0)

                _movieInputSource = retData

                movieInputSource = _movieInputSource

                _movieInputSource = Nothing

            Catch exc As Exception

                MessageBox.Show(exc.Message, " Error", MessageBoxButtons.OK)

            End Try

        End Get

    End Property

    Public ReadOnly Property movieOutputSource() As String

        Get

            'returns the current output source ex: file

            Try

                retVal = mciSendString("status movie monitor output", retData, 128, 0)

                _movieOutputSource = retData

                movieOutputSource = _movieOutputSource

                _movieOutputSource = Nothing

            Catch exc As Exception

                MessageBox.Show(exc.Message, " Error", MessageBoxButtons.OK)

            End Try

        End Get

    End Property

    Public ReadOnly Property audioStatus() As String

        Get

            'check to see if the audio is on or off

            Try

                retVal = mciSendString("status movie audio", retData, 128, 0)

                _audioStatus = retData

                audioStatus = _audioStatus

                _audioStatus = Nothing

            Catch exc As Exception

                MessageBox.Show(exc.Message, " Error", MessageBoxButtons.OK)

            End Try

        End Get

    End Property

    Public Sub sizeLocateMovie(ByVal leftPos As Integer, ByVal top As Integer, ByVal Width As Integer, ByVal Height As Integer)

        'change the size of the movie and the location of
        'the movie in Pixels

        Try
            retVal = mciSendString("put movie window at " & leftPos & " " & top & " " & Width & " " & Height, 0, 0, 0)

        Catch exc As Exception

            MessageBox.Show(exc.Message, " Error", MessageBoxButtons.OK)

        End Try

    End Sub

    Public ReadOnly Property isMoviePlaying() As Boolean

        Get

            'checks the status of the movie to see if it is playing or not

            Try

                Dim isPlaying As String = Nothing

                retVal = mciSendString("status movie mode", retData, 128, 0)

                isPlaying = Left(retData, 7)

                If isPlaying = "playing" Then

                    _isMoviePlaying = True

                Else

                    _isMoviePlaying = False

                End If

                isMoviePlaying = _isMoviePlaying

                isPlaying = Nothing
                _isMoviePlaying = Nothing

            Catch exc As Exception

                MessageBox.Show(exc.Message, " Error", MessageBoxButtons.OK)

            End Try

        End Get

    End Property

    Public ReadOnly Property returnCommandInterfaceStatus() As String

        Get

            'a very useful function for getting any errors
            'associated with the mci device. Ex: if the movie is not playing
            'this will most likely tell you the reason

            Try

                _checkError = Space(255)

                mciGetErrorString(retVal, _checkError, Len(_checkError))

                returnCommandInterfaceStatus = _checkError

                _checkError = Nothing

            Catch exc As Exception

                MessageBox.Show(exc.Message, " Error", MessageBoxButtons.OK)

            End Try

        End Get

    End Property

    Public ReadOnly Property deviceName() As String

        Get

            Try

                'returns the current device name in use

                retVal = mciSendString("info movie product", retData, 128, 0)

                _deviceName = retData

                deviceName = _deviceName

                _deviceName = Nothing

            Catch exc As Exception

                MessageBox.Show(exc.Message, " Error", MessageBoxButtons.OK)

            End Try

        End Get

    End Property

    Public ReadOnly Property deviceVersion() As String

        Get

            'returns the current version of the mci device in use

            Try

                retVal = mciSendString("info movie version", retData, 128, 0)

                _deviceVersion = retData

                deviceVersion = _deviceVersion

                _deviceVersion = Nothing

            Catch exc As Exception

                MessageBox.Show(exc.Message, " Error", MessageBoxButtons.OK)

            End Try

        End Get

    End Property

    Public ReadOnly Property nominalFrameRate() As Integer

        Get

            'returns the nominal frame rate of the movie

            Try

                retVal = mciSendString("status movie nominal frame rate wait", retData, 128, 0)

                _nominalFrameRate = Val(retData)

                nominalFrameRate = _nominalFrameRate

                _nominalFrameRate = Nothing

            Catch exc As Exception

                MessageBox.Show(exc.Message, " Error", MessageBoxButtons.OK)

            End Try

        End Get

    End Property

    Public ReadOnly Property framePerSecondRate() As Integer

        Get

            'returns the Frames Per Second of avi and mpeg movies

            Try

                retVal = mciSendString("status movie frame rate", retData, 128, 0)

                _framePerSecondRate = Val(retData) / 1000

                framePerSecondRate = _framePerSecondRate

                _framePerSecondRate = Nothing

            Catch exc As Exception

                MessageBox.Show(exc.Message, " Error", MessageBoxButtons.OK)

            End Try

        End Get

    End Property

    Public ReadOnly Property getDefaultSizeLocation() As String

        Get

            'returns the default width, height and location the movie

            Try

                retVal = mciSendString("where movie source", retData, 128, 0)

                _getDefaultSize = retData

                getDefaultSizeLocation = _getDefaultSize

                _getDefaultSize = Nothing

            Catch exc As Exception

                MessageBox.Show(exc.Message, " Error", MessageBoxButtons.OK)

            End Try


        End Get

    End Property

    Public ReadOnly Property durationInFrames() As Integer

        Get

            'get the length of the movie in frames

            Try

                retVal = mciSendString("set movie time format frames", 0, 0, 0)
                retVal = mciSendString("status movie length", retData, 128, 0)

                _durationInFrames = Val(retData)

                durationInFrames = _durationInFrames

                _durationInFrames = Nothing

            Catch exc As Exception

                MessageBox.Show(exc.Message, " Error", MessageBoxButtons.OK)

            End Try

        End Get

    End Property

    Public ReadOnly Property durationInMS() As Integer

        Get

            'get the length of the movie in milliseconds

            Try

                retVal = mciSendString("set movie time format ms", 0, 0, 0)
                retVal = mciSendString("status movie length", retData, 128, 0)

                _durationInMS = Val(retData)

                durationInMS = _durationInMS

                _durationInMS = Nothing

            Catch exc As Exception

                MessageBox.Show(exc.Message, " Error", MessageBoxButtons.OK)

            End Try

        End Get

    End Property

    Public ReadOnly Property durationInSec() As Integer

        Get

            'get the length of the movie in seconds

            Try

                _durationInSec = durationInMS / 1000

                durationInSec = _durationInSec

                _durationInSec = Nothing

            Catch exc As Exception

                MessageBox.Show(exc.Message, " Error", MessageBoxButtons.OK)

            End Try

        End Get

    End Property

    Public Sub playFullScreen()

        'play the movie in full screen mode

        Try

            retVal = mciSendString("play movie fullscreen", 0, 0, 0)

        Catch exc As Exception

            MessageBox.Show(exc.Message, " Error", MessageBoxButtons.OK)

        End Try

    End Sub

    Public WriteOnly Property shutDownVideo() As Boolean

        'false = turn off the video
        'true = turn the video on

        Set(ByVal Value As Boolean)

            Try


                If Value = False Then

                    'set the video device off
                    retVal = mciSendString("set all video off", 0, 0, 0)

                Else

                    'set the video device on
                    retVal = mciSendString("set all video on", 0, 0, 0)

                End If

            Catch exc As Exception

                MessageBox.Show(exc.Message, " Error", MessageBoxButtons.OK)

            End Try

        End Set

    End Property

    Public Sub pauseMovie()

        'pause the movie if its playing

        Try

            retVal = mciSendString("pause movie", 0, 0, 0)

        Catch exc As Exception

            MessageBox.Show(exc.Message, " Error", MessageBoxButtons.OK)

        End Try

    End Sub

    Public Function resumeMovie()

        'resumes the movie if it was paused

        Try

            retVal = mciSendString("resume movie", 0, 0, 0)

        Catch exc As Exception

            MessageBox.Show(exc.Message, " Error", MessageBoxButtons.OK)

        End Try

    End Function

    Public Property positionInFrames() As Integer

        'get or set the playing position in frames

        Get

            Try

                retVal = mciSendString("set movie time format frames wait", 0, 0, 0)
                retVal = mciSendString("status movie position", retData, 128, 0)

                _positionInFrames = Val(retData)

                positionInFrames = _positionInFrames

                _positionInFrames = Nothing

            Catch exc As Exception

                MessageBox.Show(exc.Message, " Error", MessageBoxButtons.OK)

            End Try

        End Get

        Set(ByVal Value As Integer)

            Try

                retVal = mciSendString("set movie time format frames", 0, 0, 0)

                If isMoviePlaying = True Then

                    retVal = mciSendString("play movie from " & Value, 0, 0, 0)

                Else

                    mciSendString("seek movie to " & Value, 0, 0, 0)

                End If

            Catch exc As Exception

                MessageBox.Show(exc.Message, " Error", MessageBoxButtons.OK)

            End Try

        End Set

    End Property

    Public Property positionInMS() As Integer

        Get

            'get or set the playing position of the movie in milliseconds

            Try

                retVal = mciSendString("set movie time format ms", 0, 0, 0)
                retVal = mciSendString("status movie position wait", retData, 128, 0)

                _positionInMS = Val(retData)

                positionInMS = _positionInMS

                _positionInMS = Nothing

            Catch exc As Exception

                MessageBox.Show(exc.Message, " Error", MessageBoxButtons.OK)

            End Try

        End Get

        Set(ByVal Value As Integer)

            Try

                retVal = mciSendString("set movie time format ms", 0, 0, 0)

                If isMoviePlaying() = True Then

                    mciSendString("play movie from " & Value, 0, 0, 0)

                Else

                    mciSendString("seek movie to " & Value, 0, 0, 0)

                End If

            Catch exc As Exception

                MessageBox.Show(exc.Message, " Error", MessageBoxButtons.OK)

            End Try

        End Set

    End Property

    Public Property positionInSec() As Integer

        Get

            'get or set the playing position of the movie in seconds

            Try

                _positionInSec = positionInMS / 1000

                positionInSec = _positionInSec

                _positionInSec = Nothing

            Catch exc As Exception

                MessageBox.Show(exc.Message, " Error", MessageBoxButtons.OK)

            End Try

        End Get

        Set(ByVal Value As Integer)

            Try

                Value = Value * 1000

                If isMoviePlaying() = True Then

                    mciSendString("play movie from " & Value, 0, 0, 0)

                Else

                    mciSendString("seek movie to " & Value, 0, 0, 0)

                End If

            Catch exc As Exception

                MessageBox.Show(exc.Message, " Error", MessageBoxButtons.OK)

            End Try

        End Set

    End Property

    Public Property moviePlayRate() As Integer

        Get

            'get the current speed of the movie

            Try

                retVal = mciSendString("status movie speed", retData, 128, 0)
                moviePlayRate = Val(retData)

            Catch exc As Exception

                MessageBox.Show(exc.Message, " Error", MessageBoxButtons.OK)

            End Try

        End Get

        Set(ByVal Value As Integer)

            'set the current playing speed of the movie
            '0 = as fast as possible without losing frames
            'values 1 - 2000 - 2000 being fastest. certain movie formats
            'may not support this feature

            Try

                retVal = mciSendString("set movie speed " & Value, 0, 0, 0)

            Catch exc As Exception

                MessageBox.Show(exc.Message, " Error", MessageBoxButtons.OK)

            End Try

        End Set

    End Property

    Public ReadOnly Property getPlayingStatus() As String

        Get

            'get the current mode of the movie
            'playing, stopped, paused, or not ready

            Try

                retVal = mciSendString("status movie mode", retData, 128, 0)

                _getPlayingStatus = retData 'StrConv(retData, VbStrConv.ProperCase)

                getPlayingStatus = _getPlayingStatus

                _getPlayingStatus = Nothing

            Catch exc As Exception

                MessageBox.Show(exc.Message, " Error", MessageBoxButtons.OK)

            End Try

        End Get

    End Property

    Public Sub closeMovie()

        'close the mci device. call this when you exit your application

        Try

            retVal = mciSendString("close all", 0, 0, 0)

        Catch exc As Exception

            MessageBox.Show(exc.Message, " Error", MessageBoxButtons.OK)

        End Try

    End Sub

    Public ReadOnly Property formatPosition() As String

        Get

            'get the playing position in a user friendly time format

            Try

                _formatPosition = getThisTime(positionInMS)

                formatPosition = _formatPosition

                _formatPosition = Nothing

            Catch exc As Exception

                MessageBox.Show(exc.Message, " Error", MessageBoxButtons.OK)

            End Try

        End Get

    End Property

    Public ReadOnly Property formatDuration() As String

        Get

            'get the length in a userfriendly time format

            Try

                _formatDuration = getThisTime(durationInMS)

                formatDuration = _formatDuration

                _formatDuration = Nothing

            Catch exc As Exception

                MessageBox.Show(exc.Message, " Error", MessageBoxButtons.OK)

            End Try

        End Get

    End Property

    Private Function getThisTime(ByVal timein As Integer) As String

        'used to format the position and duration of the movie

        Try

            Dim conH As Integer = Nothing
            Dim conM As Integer = Nothing
            Dim conS As Integer = Nothing
            Dim remTime As Integer = Nothing
            Dim strRetTime As String = Nothing

            remTime = timein / 1000

            conH = Int(remTime / 3600)

            remTime = remTime Mod 3600

            conM = Int(remTime / 60)

            remTime = remTime Mod 60

            conS = remTime

            If conH > 0 Then

                strRetTime = Trim(Str(conH)) & ":"

            Else

                strRetTime = ""

            End If

            If conM >= 10 Then

                strRetTime = strRetTime & Trim(Str(conM))

            ElseIf conM > 0 Then

                strRetTime = strRetTime & Trim(Str(conM))
            Else

                strRetTime = strRetTime & "0"

            End If

            strRetTime = strRetTime & ":"

            If conS >= 10 Then

                strRetTime = strRetTime & Trim(Str(conS))

            ElseIf conS > 0 Then

                strRetTime = strRetTime & "0" & Trim(Str(conS))

            Else

                strRetTime = strRetTime & "00"

            End If

            getThisTime = strRetTime

            conH = Nothing
            conM = Nothing
            conS = Nothing
            remTime = Nothing
            strRetTime = Nothing

        Catch exc As Exception

            MessageBox.Show(exc.Message, " Error", MessageBoxButtons.OK)

        End Try

    End Function

    Public Property volumeLevel() As Integer

        Get

            'get or set the current volume level
            '1000 max - 0 min

            Try

                retVal = mciSendString("status movie volume", retData, 128, 0)

                _volumeLevel = Val(retData)

                volumeLevel = _volumeLevel

                _volumeLevel = Nothing

            Catch exc As Exception

                MessageBox.Show(exc.Message, " Error", MessageBoxButtons.OK)

            End Try

        End Get

        Set(ByVal Value As Integer)

            Try

                retVal = mciSendString("setaudio movie volume to " & Value, 0, 0, 0)

            Catch exc As Exception

                MessageBox.Show(exc.Message, " Error", MessageBoxButtons.OK)

            End Try

        End Set

    End Property

    Public ReadOnly Property getVideoStatus() As String

        Get

            'get the status of the video. Returns on or off

            Try

                retVal = mciSendString("status movie video", retData, 128, 0)
                getVideoStatus = retData

            Catch exc As Exception

                MessageBox.Show(exc.Message, " Error", MessageBoxButtons.OK)

            End Try

        End Get

    End Property

    Public ReadOnly Property getCurrentTimeFormat() As String

        Get

            'returns the current time format. Should return frames or milliseconds

            Try

                retVal = mciSendString("status movie time format", retData, 128, 0)

                _getCurrentTimeFormat = retData

                getCurrentTimeFormat = _getCurrentTimeFormat

                _getCurrentTimeFormat = Nothing

            Catch exc As Exception

                MessageBox.Show(exc.Message, " Error", MessageBoxButtons.OK)

            End Try

        End Get

    End Property

    Public Property leftChannelVolume() As Integer

        Get

            'get or set the volume value of the left channel
            '1000 = max, 0 = min

            Try

                retVal = mciSendString("status movie left volume", retData, 128, 0)

                _leftChannelVolume = Val(retData)

                leftChannelVolume = _leftChannelVolume

                _leftChannelVolume = Nothing

            Catch exc As Exception

                MessageBox.Show(exc.Message, " Error", MessageBoxButtons.OK)

            End Try

        End Get

        Set(ByVal Value As Integer)

            Try

                retVal = mciSendString("setaudio movie left volume to " & Value, 0, 0, 0)

            Catch exc As Exception

                MessageBox.Show(exc.Message, " Error", MessageBoxButtons.OK)

            End Try

        End Set

    End Property

    Public Property rightChannelVolume() As Integer

        Get

            'get or set the volume value of the right channel
            '1000 = max, 0 = min

            Try

                retVal = mciSendString("status movie right volume", retData, 128, 0)

                _rightChannelVolume = Val(retData)

                rightChannelVolume = _rightChannelVolume

                _rightChannelVolume = Nothing

            Catch exc As Exception

                MessageBox.Show(exc.Message, " Error", MessageBoxButtons.OK)

            End Try

        End Get

        Set(ByVal Value As Integer)

            Try

                retVal = mciSendString("setaudio movie right volume to " & Value, 0, 0, 0)

            Catch exc As Exception

                MessageBox.Show(exc.Message, " Error", MessageBoxButtons.OK)

            End Try

        End Set

    End Property

    Public WriteOnly Property muteSoundOutput() As Boolean

        'mute/unmute both the left and right channel sound output

        Set(ByVal Value As Boolean)

            Try

                If Value = True Then

                    'turns of the audio device
                    retVal = mciSendString("set movie audio all off", 0, 0, 0)

                Else

                    'turns on the audio device
                    retVal = mciSendString("set movie audio all on", 0, 0, 0)

                End If

            Catch exc As Exception

                MessageBox.Show(exc.Message, " Error", MessageBoxButtons.OK)

            End Try

        End Set

    End Property

    Public WriteOnly Property leftChannelMute() As Boolean

        'mute/unmute just the left channel sound output

        Set(ByVal Value As Boolean)

            Try

                If Value = True Then

                    'turns of the left channel sound
                    retVal = mciSendString("set movie audio left off", 0, 0, 0)

                Else

                    'turns on the left channel sound
                    retVal = mciSendString("set movie audio left on", 0, 0, 0)

                End If

            Catch exc As Exception

                MessageBox.Show(exc.Message, " Error", MessageBoxButtons.OK)

            End Try

        End Set

    End Property

    Public WriteOnly Property rightChannelMute() As Boolean

        'mute/unmute just the right channel sound output

        Set(ByVal Value As Boolean)

            Try

                If Value = True Then

                    'turns off the right channel sound
                    retVal = mciSendString("set movie audio right off", 0, 0, 0)

                Else

                    'turns on the right channel sound
                    retVal = mciSendString("set movie audio right on", 0, 0, 0)

                End If

            Catch exc As Exception

                MessageBox.Show(exc.Message, " Error", MessageBoxButtons.OK)

            End Try

        End Set

    End Property

    Public WriteOnly Property openCdTray() As Boolean

        'true = open the cd door
        'false = close the cd door

        Set(ByVal Value As Boolean)

            Try

                If Value = True Then

                    'open the cdrom door
                    retVal = mciSendString("set cdaudio door open", 0, 0, 0)

                Else

                    'close the cdrom door
                    retVal = mciSendString("set cdaudio door closed", 0, 0, 0)

                End If

            Catch exc As Exception

                MessageBox.Show(exc.Message, " Error", MessageBoxButtons.OK)

            End Try

        End Set

    End Property

    Public Sub restartMovie()


        Try

            retVal = mciSendString("seek movie to start", 0, 0, 0)
            playMovie()

        Catch exc As Exception

            MessageBox.Show(exc.Message, " Error", MessageBoxButtons.OK)

        End Try

    End Sub

    Public WriteOnly Property rewindByMS() As Integer

        Set(ByVal Value As Integer)


            Try

                retVal = mciSendString("set movie time format ms", 0, 0, 0)
                retVal = mciSendString("play movie from " & positionInMS - Value, 0, 0, 0)

            Catch exc As Exception

                MessageBox.Show(exc.Message, " Error", MessageBoxButtons.OK)

            End Try

        End Set

    End Property

    Public WriteOnly Property rewindByFrames() As Integer

        Set(ByVal Value As Integer)

            'rewind the movie by a specified number of frames

            Try

                retVal = mciSendString("set movie time format frames", 0, 0, 0)
                retVal = mciSendString("play movie from " & positionInFrames - Value, 0, 0, 0)

            Catch exc As Exception

                MessageBox.Show(exc.Message, " Error", MessageBoxButtons.OK)

            End Try

        End Set

    End Property

    Public WriteOnly Property rewindBySeconds() As Integer

        Set(ByVal Value As Integer)

            'rewind the movie by a specified number of seconds

            Try

                retVal = mciSendString("set movie time format ms", 0, 0, 0)
                retVal = mciSendString("play movie from " & positionInMS - 1000 * Value, 0, 0, 0)

            Catch exc As Exception

                MessageBox.Show(exc.Message, " Error", MessageBoxButtons.OK)

            End Try

        End Set

    End Property

    Public WriteOnly Property forwardByFrames() As Integer

        Set(ByVal Value As Integer)

            'forward the movie a specified number of frames

            Try

                retVal = mciSendString("set movie time format frames", 0, 0, 0)
                retVal = mciSendString("play movie from " & positionInFrames + Value, 0, 0, 0)

            Catch exc As Exception

                MessageBox.Show(exc.Message, " Error", MessageBoxButtons.OK)

            End Try

        End Set

    End Property

    Public WriteOnly Property forwardByMS() As Integer

        Set(ByVal Value As Integer)

            'forward the movie a specified number of milliseconds

            Try

                retVal = mciSendString("set movie time format ms", 0, 0, 0)
                retVal = mciSendString("play movie from " & positionInMS + Value, 0, 0, 0)

            Catch exc As Exception

                MessageBox.Show(exc.Message, " Error", MessageBoxButtons.OK)

            End Try

        End Set

    End Property

    Public WriteOnly Property forwardBySeconds() As Integer

        Set(ByVal Value As Integer)

            'forward the movie a specified number of seconds

            Try

                retVal = mciSendString("set movie time format ms", 0, 0, 0)
                retVal = mciSendString("play movie from " & positionInMS + 1000 * Value, 0, 0, 0)

            Catch exc As Exception

                MessageBox.Show(exc.Message, " Error", MessageBoxButtons.OK)

            End Try

        End Set

    End Property

    Public ReadOnly Property isDeviceReady() As Boolean

        Get

            'returns true or false depending if the mci device is ready or not

            Try

                retVal = mciSendString("status movie ready", retData, 128, 0)

                _isDeviceReady = Trim(retData)

                isDeviceReady = _isDeviceReady

                _isDeviceReady = Nothing

            Catch exc As Exception

                MessageBox.Show(exc.Message, " Error", MessageBoxButtons.OK)

            End Try

        End Get

    End Property

    Public Sub timeOut(ByVal dur As Integer)

        'pauses for a specified amount of milliseconds

        Try

            Dim startTime As Integer
            Dim stopTime As Integer

            startTime = getTickCount

            Do

                Application.DoEvents()
                stopTime = getTickCount

            Loop Until stopTime - startTime >= dur

        Catch exc As Exception

            MessageBox.Show(exc.Message, " Error", MessageBoxButtons.OK)

        End Try

    End Sub

    Public Sub New()

    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
#End Region
End Class
