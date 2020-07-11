Attribute VB_Name = "NewEraFunction"
Option Explicit

'#########################################################################
'#
'#    [�V����:�ߘa]�Ή� ���t�ϊ��֐�  Ver 1.20
'#
'#       EraFormat ( �V���A���l �� ���t������    )
'#       EraCDate  ( ���t������ �� �V���A���l    )
'#       EraIsDate ( ���t�f�[�^ �� True or False )
'#
'#       Ver 0.10 , 2018/12/ 1  �b��� ���� ( EraFormat / EraCDate )
'#       Ver 0.20 , 2019/ 1/ 3  �b��� �Q��
'#
'#       Ver 1.00 , 2019/4/3
'#         (1) ������(EraFormat / EraCDate)�����[�X
'#
'#       Ver 1.10 , 2019/4/9
'#         (1) EraIsDate ��ǉ�
'#             ����ɔ���[EraCDate�̃Z������(0�`60)]������
'#
'#       Ver 1.20 , 2019/4/13
'#         (1) EraCDate��[���t]������[���l]���T�|�[�g���܂�(�V���A���l�ƌ��􂵂܂�)
'#         (2) EraCDate��[���t]������[���t������{����������]���T�|�[�g���܂�
'#             EraFormat/EraIsDate��[���t]�����ł����l�ɃT�|�[�g���܂�
'#         (3) EraCDate��[���t]�����Ő���N3��(100�`999�N),
'#             ����N2��(2000�N��Ɖ���)���T�|�[�g���܂�
'#         (4) [�ߘa]���Ή����ł�EraFormat��[���t]������
'#             [�ߘa]���t��������w��\�Ƃ��܂�
'#
'#    ���: AddinBox �p�c �j��
'#          ( http://addinbox.sakura.ne.jp/Excel_Tips28.htm )
'#
'#    -- �g�p���� --
'#    (a) EraFormat/EraCDate/EraIsDate �֐��̓t���[�E�F�A�ł��B
'#        �䎩�R�Ɋe���̃v���O�����ɑg�ݍ���ŗ��p���đՂ��č\���܂���
'#        �A����v���O�����擪�̃R�����g���K���ꏏ�� �R�s�[���Ă��������
'#
'#    (b) EraFormat/EraCDate/EraIsDate �֐���g�ݍ��񂾃v���O������
'#        �ĔЕz�ɂ������͂���܂���B
'#
'#########################################################################



'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/
'_/    [�V����:�ߘa]�Ή� ���t�ϊ��֐� EraFormat ( �V���A���l �� ���t������ )
'_/
'_/    - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'_/    �V�����ɑΉ����Ă��Ȃ��V�X�e��(Office2007�ȑO or �Ή�
'_/    �A�b�v�f�[�g���{���Ă��Ȃ�Office2010�ȍ~)�ł��A�V������
'_/    ��Â��a��ϊ����\�ɂ���֐��ł��B
'_/
'_/    Excel��TEXT�֐�/VBA��Format�֐��̑���Ɏg�p���Ă��������B
'_/    [����/�a��N]�ȊO�̕ҏW�������ꏏ�Ɏg�p���Ă���肠��܂���B
'_/
'_/    �܂��A�V�����ɑΉ��ς݂̃V�X�e���Ŏg�p���Ă���肠��܂���B
'_/
'_/    ���AEraFormat �� AddinBox/kt�֐��A�h�C��(Ver5.30)��
'_/    ktEraFormat �̖��O�Ŏ��^���܂��B
'_/    - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'_/    (��) ���t�ɃV���A���l[0�`60]�̒l���w�肵���ꍇ�ɓ�������t�ҏW�́A
'_/         �V�[�g��̏���/TEXT�֐��œ����錋�ʂƂP���Y���܂��B
'_/      EraFormat: 0��1899/12/30, 1��1899/12/31, 2��1900/1/1, 60��1900/2/28, 61��1900/3/1
'_/      �V�[�g�� : 0��1900/1/0  , 1��1900/1/1  , 2��1900/1/2, 60��1900/2/29, 61��1900/3/1
'_/      (Excel�� Lotus1-2-3�݊��ׂ̈�[1900/1/1�`1900/2/29]�̃V���A���l�������ăY�����Ă��܂�)
'_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

Public Function EraFormat(ByVal ���t As Variant, ByVal �ҏW���� As String, _
                          Optional ByVal ���N�\�L As Boolean = False) As Variant

Const cst�V�����J�n�� As Date = #5/1/2019#

Const cstEra1 As String = "�ߘa"
Const cstEra2 As String = "��"
Const cstEra3 As String = "R"

Const cstReplace_Era3 As String = "��"  '�a��N�̉I��u���p�̑�p����
Const cstReplace_ee As String = "��"
Const cstReplace_e As String = "��"
Const cstReplace_Period As String = "��"    '�s���I�h��؂�̑�p����

Dim dtm���t As Date
Dim strDateFormat As String
Dim Result As String

    '(���ӎ���)
    ' 1. LCase/UCase �ɂ�鏬����/�啶���ւ̓���ϊ��́A
    '    g/e�ȊO�̕ҏW��`���󂷂�������Ȃ��̂Ŏg��Ȃ�
    ' 2. �������p ����� �a��N�ҏW��� "0" �����l/���t�ҏW������
    '    �d������P�[�X��������邽�߂ɁAFormat�֐��̎��{�O�ł�
    '    ��p����(��,��,��)�ŉI��u�����AFormat���{��ɉ��߂Ēu������B
    ' 3. ���K�\���ɂ�錟��/�u���͗��p���܂���(VBA�ł�VBScript��K�v�Ƃ����)

    If EraIsDate(���t) Then
        '[�ߘa]���Ή����ł� [���t]������ [�ߘa]���t��������w��\�Ƃ���
        '[���N]�\�L�̓��t��������Ƃ��Ă���̂ŁA[Format�őΉ��\]����̑O��
        'EraCDate�ŃV���A���l�ϊ����s�Ȃ�
        dtm���t = EraCDate(���t)
    Else
        EraFormat = CVErr(xlErrValue)
        Exit Function
    End If

    '--- rr���� ( = gggee ) / r���� ( = ee)�̕ϊ�(Format�֐��ł͕s��) ---
    ' (rr �� r �̏��Œu������)
    ' �� LCase/UCase��g�p�̒��ӎ������Q��
    strDateFormat = Replace(�ҏW����, "rr", "gggee")
    strDateFormat = Replace(strDateFormat, "rR", "gggee")
    strDateFormat = Replace(strDateFormat, "Rr", "gggee")
    strDateFormat = Replace(strDateFormat, "RR", "gggee")
    
    strDateFormat = Replace(strDateFormat, "r", "ee")
    strDateFormat = Replace(strDateFormat, "R", "ee")



    '�ȉ��̉��ꂩ�̏����ł͖�肪�Ȃ��̂őS�� Format�֐��ɔC���Ċ����Ƃ���B
    ' (1) [����]�ȑO(2019/4/30 �ȑO)�̓��t
    ' (2) �V�����Ή��o�[�W���� or �V�����Ή��A�b�v�f�[�g���{�ς̊�
    If (dtm���t < cst�V�����J�n��) Or _
       (Format(cst�V�����J�n��, "geemmdd") = "R010501") Then
        On Error Resume Next
        Result = Format(dtm���t, strDateFormat)    ' Format�֐��ɂ��ҏW
        If (Err.Number <> 0) Then
            EraFormat = CVErr(xlErrValue)
            Exit Function
        End If
        On Error GoTo 0
        ' ���N��1�N �ҏW
        EraFormat = prvFirstYearEdit(���N�\�L, dtm���t, Result, �ҏW����)
        Exit Function
    End If



    '####################################################################
    '###                                                              ###
    '###  �ȍ~ [���t��2019/5/1 �� [�V����]���Ή���] ����̕ϊ�����  ###
    '###                                                              ###
    '####################################################################

    '���P�[��ID���L��Ύ�菜��([$-411] : ���{ , [$-409] : �č�)
    'Format�֐��ł̓��P�[��ID�������Ă����Ғʂ�ɕϊ������
    strDateFormat = Replace(strDateFormat, "[$-411]", "")
    strDateFormat = Replace(strDateFormat, "[$-409]", "")
    If (strDateFormat = "") Then
        EraFormat = CVErr(xlErrValue)
        Exit Function
    End If
    
    '--- ggg��"�ߘa" , gg��"��" , g��"R"(��p������) �ɕϊ����� ---
    ' (ggg �� gg �� g �̏��Œu������)
    ' �� LCase/UCase��g�p�̒��ӎ������Q��
    strDateFormat = Replace(strDateFormat, "ggg", cstEra1)
    strDateFormat = Replace(strDateFormat, "Ggg", cstEra1)
    strDateFormat = Replace(strDateFormat, "gGg", cstEra1)
    strDateFormat = Replace(strDateFormat, "ggG", cstEra1)
    strDateFormat = Replace(strDateFormat, "GGg", cstEra1)
    strDateFormat = Replace(strDateFormat, "GgG", cstEra1)
    strDateFormat = Replace(strDateFormat, "gGG", cstEra1)
    strDateFormat = Replace(strDateFormat, "GGG", cstEra1)
    
    strDateFormat = Replace(strDateFormat, "gg", cstEra2)
    strDateFormat = Replace(strDateFormat, "gG", cstEra2)
    strDateFormat = Replace(strDateFormat, "Gg", cstEra2)
    strDateFormat = Replace(strDateFormat, "GG", cstEra2)
    
    strDateFormat = Replace(strDateFormat, "g", cstReplace_Era3)
    strDateFormat = Replace(strDateFormat, "G", cstReplace_Era3)
    
    '--- ee / e ��a��N(����N - 2018) �ɕϊ�����(��p���� ��,��) ---
    ' �� �����܂ŏ���������ė���̂�2019/5/1�ȍ~�̓��t�̂݁B
    '    �����ȑO(2019/4/30�ȑO)�̓��t�́A�����܂ŗ���ė��Ȃ��̂�
    '    �a��N�̊��Z����[����N-2018]�Œ�ő��v
    ' (��) [�a��N]�ҏW������ e / ee �̂݁Beee �͖����B
    '      eee�N�� [ee]+[e�N]�Ɖ���(1�N�� 011�N �ɂȂ�)�����B
    
    ' (ee �� e �̏��Œu������)
    strDateFormat = Replace(strDateFormat, "ee", cstReplace_ee)
    strDateFormat = Replace(strDateFormat, "eE", cstReplace_ee)
    strDateFormat = Replace(strDateFormat, "Ee", cstReplace_ee)
    strDateFormat = Replace(strDateFormat, "EE", cstReplace_ee)
    
    strDateFormat = Replace(strDateFormat, "e", cstReplace_e)
    strDateFormat = Replace(strDateFormat, "E", cstReplace_e)
    
    
    '---�y g , e �ȊO�̕ҏW�� Format�֐��ɔC���� �z---
    '
    ' �A���AFormmat �ɔC����O�� �s���I�h����p�����ɒu�����Ă����K�v������B
    ' ���R�F"ge.m.d"��"����.m.d" �ƂȂ邪�A[�N]�ҏW�����������Ȃ�ׂ�
    '       �ŏ��̃s���I�h�������_�Ɖ��߂���Ă��܂��܂��B
    '       ���̌��ʁA�V���A���l���s���I�h�ʒu�ɐ��l�Ƃ��ĕ\������܂��B
    '       �c��� "m.d" �������ҏW�����ł͂Ȃ��P�Ȃ�Œ�\�������Ƃ���
    '       ���߂���āAm.d �̕����ł��̂܂ܕ\������܂��B
    strDateFormat = Replace(strDateFormat, ".", cstReplace_Period)
    
    On Error Resume Next
    Result = Format(dtm���t, strDateFormat)    ' Format�֐��ɂ��ҏW
    If (Err.Number <> 0) Then
        EraFormat = CVErr(xlErrValue)
        Exit Function
    End If
    On Error GoTo 0
    '----------------------------------------------------

    ' �I��u���̑�p����(��,��,��,��)��
    ' �{���̒l(�������p, �a��N2��, �a��N1��, �s���I�h)����������
    ' �� �����܂ŏ���������ė���̂�2019/5/1�ȍ~�̓��t�̂݁B
    '    �����ȑO(2019/4/30�ȑO)�̓��t�́A�����܂ŗ���ė��Ȃ��̂�
    '    �a��N�̊��Z����[����N-2018]�Œ�ő��v
    Result = Replace(Result, cstReplace_Era3, cstEra3)
    Result = Replace(Result, cstReplace_ee, Format(Year(dtm���t) - 2018, "00"))
    Result = Replace(Result, cstReplace_e, Format(Year(dtm���t) - 2018, "0"))
    Result = Replace(Result, cstReplace_Period, ".")


    ' ���N��1�N �ҏW
    EraFormat = prvFirstYearEdit(���N�\�L, dtm���t, Result, �ҏW����)

End Function

'-----------------------------------------------------------------------------
' "����01�N"/"����1�N"/"��01�N"/"��1�N"/"H01�N"/"H1�N" ����"���N"�\�L�ɉ��߂܂��B
'
' (��) �����֘A�̃A�b�v�f�[�g��VBA��Format�֐����̂Ɂu���N�v�\�L�̋@�\���ǉ�����܂����B
'      (���W�X�g��(InitialEraYear)��[���N]�w�肪�K�v)
'                              -- �A�b�v�f�[�g�ϊ� , ���A�b�v�f�[�g��
'   Format("2019/5/1","ggge�N") ��   "�ߘa���N"      ,  "�ߘa1�N
'   Format("2019/5/1","gge�N")  ��   "�ߌ��N"        ,  "��1�N"
'   Format("2019/5/1","ge�N")   ��   "R���N"         ,  "R1�N"
'
'  Format�֐��ɂ�����"���N"�ƕҏW����Ă���P�[�X�ł́A
'  EraFormat�֐���[���N�\�L]�����̎w��ɍ��킹�āA
'  False(���N�\�L�Ȃ�)�̏ꍇ�ɂ� "���N"��"1�N" or "01�N" �ɖ߂��܂��B
'-----------------------------------------------------------------------------
Private Function prvFirstYearEdit(ByVal FirstYear As Boolean, _
                                  ByVal SerialDate As Date, _
                                  ByVal EditDate As String, _
                                  ByVal EditPattern As String) As String
Dim strEdit As String

    If (LCase(EditPattern) Like "*e�N*") Then
        ' [�a��+"�N"]�ҏW����
    Else
        ' [�a��+"�N"]�ҏW�Ȃ� �c [1�N�̌��N]�ϊ������͕s�v
        prvFirstYearEdit = EditDate
        Exit Function
    End If

    If (FirstYear = False) Then
        ' "�ߘa���N"��"�ߘa1�N" ���A"1�N"�\�L�ɖ߂�
        ' ����"���N"�\�L�ɂȂ��Ă�����̂Ȃ̂Ō����܂Ń`�F�b�N����K�v�Ȃ�
        Select Case Year(SerialDate)
          Case 1868, 1912, 1926, 1989, 2019
            ' ����1�N(1868), �吳1�N(1912), ���a1�N(1926), ����1�N(1989), �ߘa1�N(2019)
            strEdit = prvFormat_GannenTo1st(EditDate, EditPattern)
            
          Case Else
            '�����ȑO or �e������[�Q�N�`]
            strEdit = EditDate
        End Select
    Else
        ' "�ߘa1�N"��"�ߘa���N" ���A"���N"�\�L�ɉ��߂�
        Select Case Format(SerialDate, "yyyymmdd")
          Case "18681023" To "18681231"     '���� ���N
            strEdit = prvFormat_1stToGannen(EditDate)
        
          Case "19120730" To "19121231"     '�吳 ���N
            strEdit = prvFormat_1stToGannen(EditDate)
        
          Case "19261225" To "19261231"     '���a ���N
            strEdit = prvFormat_1stToGannen(EditDate)
        
          Case "19890108" To "19891231"     '���� ���N
            strEdit = prvFormat_1stToGannen(EditDate)
        
         Case "20190501" To "20191231"     '�ߘa ���N
            strEdit = prvFormat_1stToGannen(EditDate)
        
         Case Else
            '�����ȑO or �e������[�Q�N�`]
            strEdit = EditDate
       End Select
    End If
    
    prvFirstYearEdit = strEdit
End Function

'-----------------------------------------------------------------------------
' [���N��1�N] for EraFormat
'-----------------------------------------------------------------------------
' (a) �ҏW�����ɂ� ee�N or e�N ��[�����Ȃ�]�̃p�^�[�����L�蓾��̂�
'     �^�[�Q�b�g��[����+"���N"]���ɂ��Ă͑ʖ�
' (b) ee�N��"01�N" , e�N��"1�N"
'
' (��) [�a��N]�ҏW������ e / ee �̂݁Beee �͖����B
'      eee�N�� [ee]+[e�N]�Ɖ���(1�N�� 011�N �ɂȂ�)�����B
'-----------------------------------------------------------------------------
Private Function prvFormat_GannenTo1st(ByVal EditDate As String, _
                                       ByVal EditPattern As String) As String
Dim strEdit As String

    If (LCase(EditPattern) Like "*ee�N*") Then
        strEdit = Replace(EditDate, "���N", "01�N")
    Else
        strEdit = Replace(EditDate, "���N", "1�N")
    End If
    
    prvFormat_GannenTo1st = strEdit
End Function

'-----------------------------------------------------------------------------
' [01�N or 1�N�ˌ��N] for EraFormat
'-----------------------------------------------------------------------------
' (a) �ҏW�����ɂ� ee�N or e�N ��[�����Ȃ�]�̃p�^�[�����L�蓾��̂�
'     �^�[�Q�b�g��[����+1�N]���ɂ��Ă͑ʖ�
' (b) "yyyy�N(gggee�N)"�̂悤��[����N]�ƕ��L�̃p�^�[���ł�
'     ���̊֐����Ăяo�����̂͌��N(1868,1912,1926,1989,2019)�̏ꍇ�����Ȃ̂�
'     [����N]������ "01�N"/"1�N"�ɂȂ邱�Ƃ͂Ȃ��B
'
' (��) [�a��N]�ҏW������ e / ee �̂݁Beee �͖����̂ŁA"001�N"�̕ϊ��͕s�v�B
'      eee�N�� [ee]+[e�N]�Ɖ���(1�N�� 011�N �ɂȂ�)�����B
'-----------------------------------------------------------------------------
Private Function prvFormat_1stToGannen(ByVal EditDate As String) As String
Dim strEdit As String

    ' "01�N"�̒u�� �� "1�N"�̒u�� �̏��ōs�Ȃ���
    strEdit = Replace(EditDate, "01�N", "���N")
    strEdit = Replace(strEdit, "1�N", "���N")
    
    prvFormat_1stToGannen = strEdit
End Function


'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/
'_/    [�V����:�ߘa]�Ή� ���t�ϊ��֐� EraCDate    ( ���t������ �� �V���A���l )
'_/
'_/    - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'_/    �V�����ɑΉ����Ă��Ȃ��V�X�e��(Office2007�ȑO or �Ή�
'_/    �A�b�v�f�[�g���{���Ă��Ȃ�Office2010�ȍ~)�ł��A�V������
'_/    ��Â��a����t����t�f�[�^(�V���A���l)�ɕϊ��ł���֐��ł��B
'_/
'_/    Excel��DATEVALUE�֐�/VBA��CDate/DateValue�֐��̑���Ɏg�p���Ă��������B
'_/
'_/    �V�����ɑΉ��ς݂̃V�X�e���Ŏg�p���Ă���肠��܂���B
'_/
'_/    ���AEraCDate �� AddinBox/kt�֐��A�h�C��(Ver5.30)��
'_/    ktEraCDate �̖��O�Ŏ��^���܂��B
'_/    - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'_/   �y EraCDate �ŃT�|�[�g������t������̃t�H�[�}�b�g �z
'_/      a) ��؂�`��(�a��N��1�`3��,4���ȏ�̓G���[)
'_/           [H31�N4��30��] [H31/4/30] [H31-4-30] [H31.4.30]
'_/      b) ����
'_/           [����,��,M/m] [�吳,��,T/t] [���a,��,S/s] [����,��,H/h] [�ߘa,��,R/r]
'_/      c) ����N(4��,3��,2��)���ϊ��\�ł�(2����2000�N��Ɖ��߂��܂�)
'_/           [2019�N4��30��] [2019/4/30] [2019-4-30] [2019.4.30]
'_/         ���A���m����[��/��/�N or ��/��/�N]�t�H�[�}�b�g��NG�ł�
'_/      d) ����32 ���̉����Ȍ�̔N���ł��n�j�Ƃ��Ă��܂�
'_/         �A���A[�����͈�=True]�w��̏ꍇ�͌������ԓ��̓��t�݂̂� OK �ƂȂ�܂��B
'_/      e) "�������N","�吳���N","���a���N","�������N","�ߘa���N"�Ƃ����\�L���Ƃ��܂��B
'_/      f) ���������񂪑����Ă���ꍇ�A�������݂ŕϊ����܂��B
'_/         �A���A���̎���������VBA��CDate�֐��ŕϊ��\�ȃt�H�[�}�b�g�Ɍ���܂��B
'_/
'_/    (��) ���t�������[1900/1/1�`1900/2/29]�̊��Ԃ̓��t���w�肵���ꍇ�ɓ�����
'_/         �V���A���l�́A�V�[�g��ɓ��͂����ꍇ�̒l�Ƃ͂P���Y���܂��B
'_/      EraCDate : 1900/1/1��2, 1900/2/28��60, 1900/2/29��#VALUE!, 1900/3/1��61
'_/      �V�[�g�� : 1900/1/1��1, 1900/2/28��59, 1900/2/29��60     , 1900/3/1��61
'_/      (Excel�� Lotus1-2-3�݊��ׂ̈�[1900/1/1�`1900/2/29]�̃V���A���l�������ăY�����Ă��܂�)
'_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

Public Function EraCDate(ByVal ���t������ As String, _
                         Optional ByVal �����͈� As Boolean = False) As Variant

'�� ����[���t������]��Date�^�f�[�^���w�肵���ꍇ�A
'   ���t������ϊ�����āy"yyyy/m/d"�`���̓��t������z�Ƃ��Ď󂯎��B
'   Date�^�f�[�^�F[Date�^�̕ϐ�][DateValue/DateSerial/CDate�̌���][���t�����̃Z���l]
'   [���t�����̃Z���l]�͘a����ł����Ă� "yyyy/m/d"�`���ƂȂ��Ď󂯎��
'
'�� ����[���t������]�ɐ��l or ������������w�肵���ꍇ�A
'   Date�^�͈̔� [100/1/1(-657434)�`9999/12/31(2958465)] ���ł���΃V���A���l�Ƃ��ĕԂ��B
'
'�� ����[���t������]�Ɂu��Z���v���w�肵���ꍇ�́u�󕶎��v�Ƃ��Ď󂯎��̂ŃG���[�ɂȂ�B

Dim strDate As String
Dim dtmDate As Date
Dim strTemp As String
Dim intEra As Integer   ' 0:����, 1:����, 2:�吳, 3:���a, 4:����, 5:�ߘa
Dim aryEraRange As Variant

Const cstPattern As String = "[�������Ö��叺����MTSHRmtshr]"

    ' "�������N","�����N","H���N" ����"���N"��"1�N"�\�L�ɉ��߂�
    ' �������t���Ȃ�"���N"�݂͔̂N�オ����ł��Ȃ��ׁA�ϊ��G���[�ł�
    strTemp = prvCDate_GannenTo1st(���t������, "����", "��", "M")
    strTemp = prvCDate_GannenTo1st(strTemp, "�吳", "��", "T")
    strTemp = prvCDate_GannenTo1st(strTemp, "���a", "��", "S")
    strTemp = prvCDate_GannenTo1st(strTemp, "����", "��", "H")
    strTemp = prvCDate_GannenTo1st(strTemp, "�ߘa", "��", "R")

    ' ����(�Q����)�̔��菈���� Like ���Z�q�ōs�Ȃ���l�ɑ�֕���(�P����)�Œu������
    strTemp = Replace(strTemp, "����", "��")
    strTemp = Replace(strTemp, "�吳", "��")
    strTemp = Replace(strTemp, "���a", "��")
    strTemp = Replace(strTemp, "����", "��")
    strTemp = Replace(strTemp, "�ߘa", "��")

    '=== ���t������̃p�^�[���`�F�b�N ===
    '=== �N�͂����Ő������������(# �` #### �p�^�[��)����
    '=== �����̐��������CDate�ɔC����(�Œ��,1�������͕ۏ�)
    '=== [���t������{����������]���ΏۂƂ���(������*�Ŏ��������� �����L���Ă��n�j�ɂȂ�)

    '--- �a��(�N1��) ---
    If (strTemp Like cstPattern & "#�N#*��#*��*") Or _
       (strTemp Like cstPattern & "#/#*/#*") Or _
       (strTemp Like cstPattern & "#.#*.#*") Or _
       (strTemp Like cstPattern & "#-#*-#*") Then
        If (Mid(strTemp, 2, 1) = "0") Then
            EraCDate = CVErr(xlErrValue)  ' 0�N �̓G���[
            Exit Function
        Else
            '[����+�a��N]��[����N]�ϊ�
            strDate = prvEraYear4EraCDate(Mid(strTemp, 1, 1), Mid(strTemp, 2, 1), Mid(strTemp, 3), intEra)
        End If

    '--- �a��(�N2��) ---
    ElseIf (strTemp Like cstPattern & "##�N#*��#*��*") Or _
           (strTemp Like cstPattern & "##/#*/#*") Or _
           (strTemp Like cstPattern & "##.#*.#*") Or _
           (strTemp Like cstPattern & "##-#*-#*") Then
        If (Mid(strTemp, 2, 2) = "00") Then
            EraCDate = CVErr(xlErrValue)    ' 00�N �̓G���[
            Exit Function
        Else
            '[����+�a��N]��[����N]�ϊ�
            strDate = prvEraYear4EraCDate(Mid(strTemp, 1, 1), Mid(strTemp, 2, 2), Mid(strTemp, 4), intEra)
        End If

    ' (��) [�a��N]�ҏW������ e / ee �̂݁Beee �͖����̂ŁA�a��(�N3��)�̔���͕s�v�B
    '      eee�N�� [ee]+[e�N]�Ɖ���(1�N�� 011�N �ɂȂ�)�����B

    '--- ����(�N4��) ---
    ' ���m����[��/��/�N or ��/��/�N]�t�H�[�}�b�g��NG�ł�
    ElseIf (���t������ Like "####�N#*��#*��*") Or _
           (���t������ Like "####/#*/#*") Or _
           (���t������ Like "####.#*.#*") Or _
           (���t������ Like "####-#*-#*") Then
        '�s���I�h��؂肪DateValue�ł͕ϊ��ł��Ȃ��̂� / �ɒu������
        strDate = Replace(���t������, ".", "/")
        intEra = 0

    '--- ����(�N3��) --- (���̂܂� 100�`999�N�Ɖ��߂���)
    ElseIf (���t������ Like "###�N#*��#*��*") Or _
           (���t������ Like "###/#*/#*") Or _
           (���t������ Like "###.#*.#*") Or _
           (���t������ Like "###-#*-#*") Then
        '�s���I�h��؂肪DateValue�ł͕ϊ��ł��Ȃ��̂� / �ɒu������
        strDate = Replace(���t������, ".", "/")
        intEra = 0

    '--- ����(�N2��) --- (2000�N��Ɖ��߂���)
    ElseIf (���t������ Like "##�N#*��#*��*") Or _
           (���t������ Like "##/#*/#*") Or _
           (���t������ Like "##.#*.#*") Or _
           (���t������ Like "##-#*-#*") Then
        '�s���I�h��؂肪DateValue�ł͕ϊ��ł��Ȃ��̂� / �ɒu������
        strDate = "20" & Replace(���t������, ".", "/")  '�擪��"20"��t������2000�N��
        intEra = 0

    '--- ���l or ���������� ---
    ' �����ŁA�� CDate�ϊ�(�P�Ȃ�Date�^�ւ̌^�ϊ�)���Ēl��Ԃ�
    ElseIf IsNumeric(���t������) Then
        '�� �J���}�ҏW���l�̏ꍇ�̓J���}����菜��
        '   �J���}�����t��؂蕶���Ƃ��Ĉ����Ă��܂��A
        '   �ȗ��`�̓��t������Ƃ��ė\�z�O�̓��t�ƌ��􂳂��
        '   CDate("2,500")��500/2/1 �Ɖ���(m,yyy)�����
        On Error Resume Next
        dtmDate = CDate(Replace(���t������, ",", ""))
        If (Err.Number <> 0) Then
            '�V���A���l�͈͊O [100/1/1(-657434)�`9999/12/31(2958465)]
            EraCDate = CVErr(xlErrValue)
        Else
            EraCDate = dtmDate
        End If
        On Error GoTo 0
        Exit Function

    '--- ���̓G���[ ---
    Else
        EraCDate = CVErr(xlErrValue)
        Exit Function
    End If

    '=== CDate�ɂ����t������(������������܂�ł���)�˃V���A���l �ϊ� ===
    ' �a��N�͐���N�ɕϊ���
    On Error Resume Next
    dtmDate = CDate(strDate)
    If (Err.Number <> 0) Then
        EraCDate = CVErr(xlErrValue)
        Exit Function
    End If
    On Error GoTo 0

    If (�����͈� = True) Then
        If (intEra = 0) Then
            '����͔͈͖���
        Else
            ' ����[1]�F1868(M1)/10/23 �` 1912(M45)/ 7/29
            ' �吳[2]�F1912(T1)/ 7/30 �` 1926(T15)/12/24
            ' ���a[3]�F1926(S1)/12/25 �` 1989(S64)/ 1/ 7
            ' ����[4]�F1989(H1)/ 1/ 8 �` 2019(H31)/ 4/30
            ' �ߘa[5]�F2019(R1)/ 5/ 1 �` 9999(---)/12/31
            ' ��[����]���w�肳��Ă���ꍇ�ɂ͏���������̂ŁA
            '   �I�[�̔����[�I�����{�P��菬]�Ƃ���
            aryEraRange = _
                Array(Array(0, 0), Array(#10/23/1868#, #7/29/1912#), _
                      Array(#7/30/1912#, #12/24/1926#), Array(#12/25/1926#, #1/7/1989#), _
                      Array(#1/8/1989#, #4/30/2019#), Array(#5/1/2019#, #12/31/9999#))
            If (aryEraRange(intEra)(0) <= dtmDate) And _
               ((aryEraRange(intEra)(1) + 1) > dtmDate) Then
                '�����͈͓��łn�j
            Else
                EraCDate = CVErr(xlErrValue)
                Exit Function
            End If
        End If
    End If

    EraCDate = dtmDate

End Function

'-----------------------------------------------------------------------------
' [���N��1�N] for EraCDate   Era1:�ߘa, Era2:��, Era3:R ��
'-----------------------------------------------------------------------------
Private Function prvCDate_GannenTo1st _
            (ByVal EditDate As String, ByVal Era1 As String, _
             ByVal Era2 As String, ByVal Era3 As String) As String
Dim strEdit As String

    strEdit = Replace(EditDate, (Era1 & "���N"), (Era1 & "1�N"))
    strEdit = Replace(strEdit, (Era2 & "���N"), (Era2 & "1�N"))
    strEdit = Replace(strEdit, (Era3 & "���N"), (Era3 & "1�N"))
    strEdit = Replace(strEdit, (LCase(Era3) & "���N"), (Era3 & "1�N"))
    
    prvCDate_GannenTo1st = strEdit
End Function

'-----------------------------------------------------------------------------
' [����+�a��N] �� [����N]�ϊ� , ������[�����t���O]�ŕԂ�
'
' �Q���������̓�(����),��(�吳),��(���a),��(����),��(�ߘa)�ɒu������Ă���
' [�N]�� 1�`2�����̐���
' [����]�ɂ�[�N]�����ȍ~�̓��t������(������������܂�)�̕������n�����
'-----------------------------------------------------------------------------
Private Function prvEraYear4EraCDate(ByVal ���� As String, ByVal �N As String, _
                                     ByVal ���� As String, ByRef �����t���O As Integer) As String
Dim strDate As String
Dim strMMDD As String

    '�s���I�h��؂肪DateValue�ł͕ϊ��ł��Ȃ��ׁA�X���b�V���ɒu������
    strMMDD = Replace(����, ".", "/")

    '�����ɉ����Đ���N�ɕϊ�����(�N�� Like���Z�ɂ�萔���̃`�F�b�N��)
    Select Case ����
      Case "��", "��", "M", "m"
        strDate = (CLng(�N) + 1867) & strMMDD  '����
        �����t���O = 1
      Case "��", "��", "T", "t"
        strDate = (CLng(�N) + 1911) & strMMDD  '�吳
        �����t���O = 2
      Case "��", "��", "S", "s"
        strDate = (CLng(�N) + 1925) & strMMDD  '���a
        �����t���O = 3
      Case "��", "��", "H", "h"
        strDate = (CLng(�N) + 1988) & strMMDD  '����
        �����t���O = 4
      Case "��", "��", "R", "r"
        strDate = (CLng(�N) + 2018) & strMMDD  '�ߘa
        �����t���O = 5
    End Select
    prvEraYear4EraCDate = strDate
End Function


'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/
'_/    [�V����:�ߘa]�Ή� ���t����֐� EraIsDate    ( ���t�f�[�^ �� True/False )
'_/
'_/    �V�����ɑΉ����Ă��Ȃ��V�X�e��(Office2007�ȑO or �Ή�
'_/    �A�b�v�f�[�g���{���Ă��Ȃ�Office2010�ȍ~)�ł��A�V������
'_/    ��Â��a����t���܂߁A�u���t�Ƃ��đÓ����ۂ��v�𔻒肷��֐��ł��B
'_/
'_/    EraIsDate ���T�|�[�g������t������̃t�H�[�}�b�g��EraCDate�ɏ����܂��B
'_/
'_/    �V�����ɑΉ��ς݂̃V�X�e���Ŏg�p���Ă���肠��܂���B
'_/
'_/    ���AEraIsDate �� AddinBox/kt�֐��A�h�C��(Ver5.30)��
'_/    ktEraIsDate �̖��O�Ŏ��^���܂��B
'_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

Public Function EraIsDate(ByVal ���t�f�[�^ As Variant) As Boolean
Dim strData As String
Dim Result As Variant

    If IsError(���t�f�[�^) Then
        EraIsDate = False

    ElseIf IsEmpty(���t�f�[�^) Then  '��Z����Empty�l
        EraIsDate = False

    ElseIf IsDate(���t�f�[�^) Then
        'IsDate�̔���ΏہF�u���t������v�uDate�^�̃f�[�^�v
        '(1) EraCDate�ł�IsDate���肪�\�����A�����̖ʂ���
        '    IsDate�őΉ��\�ȕ����͐��IsDate�ōς܂���
        '
        '(2) [IsDate(Date�^�f�[�^) �� True]
        '
        '(3) [IsDate("����33�N2��29��") ��False]
        '    �V�[�g��ł̃V���A���l60�ɑ΂�����t������
        '    EraCDate�ł� False �ɂȂ�
        '
        '(4) [IsDate("���a65�N1��1��")��False (�����͈̓I�[�o�[)] �c ���EraCDate�ŋ~��
        '    [IsDate(���l)��False] �c ���� IsNumeric�ŋ~��
        '
        '(5) VBA�Ȃ̂Ń}�C�i�X�̃V���A���l�ɑΉ�����
        '    ���t������("M1/10/23" , "1899/1/1" ��)�� True �ɂȂ�
        '
        '(6) [�ߘa]���t������
        '      �V�����A�b�v�f�[�g�� �� True
        '      �V�����A�b�v�f�[�g�� �� False �c ���EraCDate�ŋ~��
        '
        '(7) "�������N1��8��"����[���N�\�L]
        '      �V�����A�b�v�f�[�g�� �� True
        '      �V�����A�b�v�f�[�g�� �� False �c ���EraCDate�ŋ~��

        EraIsDate = True

    ElseIf IsNumeric(���t�f�[�^) Then
        'IsNumeric�̔���ΏہF�u���l , ����������v
        '(1) �P�Ȃ鐔�l����Ȃ̂Ń}�C�i�X�l(1899�N�ȑO�̓��t)�� True �ɂȂ�
        '
        '(2) �V���A���l�͈̔͂� 100/1/1(-657434)�`9999/12/31(2958465)
        
        If (VarType(���t�f�[�^) = vbString) Then
            If (CDbl(���t�f�[�^) >= -657434) And (CDbl(���t�f�[�^) <= 2958465) Then
                EraIsDate = True
            Else
                EraIsDate = False
            End If
        ElseIf (���t�f�[�^ >= -657434) And (���t�f�[�^ <= 2958465) Then
            EraIsDate = True
        Else
            EraIsDate = False
        End If

    Else
        'Else�̔���ΏہF
        '  ���l/������ȊO�̃f�[�^�^��False
        '  IsDate �Œe���ꂽ���t������(���L)�̋~�ρ�True
        '(1) [�ߘa]���t������
        '(2) "�������N1��8��"����[���N�\�L]
        '(3) "���a65�N1��1��"���̌����͈̓I�[�o�[
        On Error Resume Next
        strData = CStr(���t�f�[�^)
        If (Err.Number <> 0) Then
            EraIsDate = False   '���l/������ȊO�̃f�[�^�͑ΏۊO
            Exit Function
        End If
        On Error GoTo 0
        
        Result = EraCDate(strData, False)   'EraCDate�Ō��؂���(�����͈͖͂���)
        If IsError(Result) Then
            EraIsDate = False
        Else
            EraIsDate = True
        End If
    End If
End Function

