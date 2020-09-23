Attribute VB_Name = "SoundsMod"
Option Explicit
Private m_snd() As Byte
Private Const SND_SYNC = &H0
Private Const SND_ASYNC = &H1
Private Const SND_NODEFAULT = &H2
Private Const SND_MEMORY = &H4
Private Const SND_ALIAS = &H10000
Private Const SND_FILENAME = &H20000
Private Const SND_RESOURCE = &H40004
Private Const SND_ALIAS_ID = &H110000
Private Const SND_ALIAS_START = 0
Private Const SND_LOOP = &H8
Private Const SND_NOSTOP = &H10
Private Const SND_VALID = &H1F
Private Const SND_NOWAIT = &H2000
Private Const SND_VALIDFLAGS = &H17201F
Private Const SND_RESERVED = &HFF000000
Private Const SND_TYPE_MASK = &H170007
Private Const WAVERR_BASE = 32
Private Const WAVERR_BADFORMAT = (WAVERR_BASE + 0)
Private Const WAVERR_STILLPLAYING = (WAVERR_BASE + 1)
Private Const WAVERR_UNPREPARED = (WAVERR_BASE + 2)
Private Const WAVERR_SYNC = (WAVERR_BASE + 3)
Private Const WAVERR_LASTERROR = (WAVERR_BASE + 3)
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Private Declare Function PlaySoundData Lib "winmm.dll" Alias "PlaySoundA" (lpData As Any, ByVal hModule As Long, ByVal dwFlags As Long) As Long

Public Function PlaySoundResource(ByVal SndID As Long) As Long
   Const Flags = SND_MEMORY Or SND_ASYNC Or SND_NODEFAULT
   m_snd = LoadResData(SndID, "SOUND")
   PlaySoundResource = PlaySoundData(m_snd(0), 0, Flags)
End Function

