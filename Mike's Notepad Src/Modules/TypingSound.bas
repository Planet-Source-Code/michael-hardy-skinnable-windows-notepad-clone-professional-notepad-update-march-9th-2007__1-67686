Attribute VB_Name = "TypingSound"
Option Explicit
Public Const Typer = 101
Private m_snd() As Byte
Private Const SND_ASYNC = &H1
Private Const SND_MEMORY = &H4
Private Const SND_LOOP = &H8
Private Declare Function PlaySoundData Lib "winmm.dll" Alias "PlaySoundA" _
(lpData As Any, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Public Function PlaySound(ByVal SndID As Long) As Long
Const Flags = SND_ASYNC Or SND_MEMORY
m_snd = LoadResData(SndID, "SOUND")
PlaySoundData m_snd(0), 0, Flags
End Function
Public Function PlaySoundloop(ByVal SndID As Long) As Long
Const Flags = SND_ASYNC Or SND_MEMORY Or SND_LOOP
m_snd = LoadResData(SndID, "SOUND")
PlaySoundData m_snd(0), 0, Flags
End Function



