Attribute VB_Name = "basImageCount"
Option Explicit

Sub frmImageCount_Show(ByVal control As IRibbonControl)
    frmImageCount.Show
End Sub

Sub Check_Image_Type_Powered_By_ImageMagick(Image_Files As Variant, Page_flg As Boolean, Compress_flg As Boolean, Pixel_flg As Boolean, DPI_flg As Boolean)
    Dim Image_File As Variant                       '�P�̂̃t�@�C���p�X
    Dim File_Count As Long                          '�i���\���p�̑I���t�@�C�������i�[
    Dim i As Long                                   '�t�@�C�����Ƃ̃C���f�b�N�X
    Dim j As Long                                   '���ڂ��Ƃ̃C���f�b�N�X
    Dim img As Object                               'ImageMagick�̃I�u�W�F�N�g
    Dim Headder_Dic As New Scripting.Dictionary     '�w�b�_�[���ڂ̗�Ԓn�i�[�p
    Dim c As Variant                                '���ڂ�key�i�[
    Dim objpb                                       '�i���o�[�̃I�u�W�F�N�g
    
    '�`��X�V���Ȃ�
    Application.ScreenUpdating = False
    
    '�摜�ȊO�̂��̂��o�Ă��Ă��G���[�Ŏ~�܂�Ȃ��悤�ɂ���
    On Error Resume Next
    
    '�i���\���p�̃t�@�C���I�𐔂��i�[
    File_Count = UBound(Image_Files)
    
    'ImageMagick���g�����߂ɃI�u�W�F�N�g���쐬
    Set img = New ImageMagickObject.MagickImage
    
    '�V�[�g��V�K�쐬���Ĉ�Ԍ��ɔz�u
    Worksheets.Add after:=Worksheets(Worksheets.Count)
    
    '�W�J����f�[�^�̊J�n�s
    i = 1
    
    '���ڃC���f�b�N�X�̏�����
    j = 1
    
    '�e�`�F�b�N�{�b�N�X�̃t���O����A�z�z��ɕϊ�����
    
    '���ꂾ���͌Œ�
    Headder_Dic.Add "�t�@�C���p�X", j
    
    '��������t���O����A�z�z��ɕϊ�
    If Page_flg Then
        Headder_Dic.Add "�y�[�W��", j + 1
        j = j + 1
    End If
    
    If Compress_flg Then
        Headder_Dic.Add "���k�`��", j + 1
        j = j + 1
    End If
    
    If Pixel_flg Then
        Headder_Dic.Add "�s�N�Z����_�c", j + 1
        Headder_Dic.Add "�s�N�Z����_��", j + 2
        j = j + 2
    End If
    
    If DPI_flg Then
        Headder_Dic.Add "DPI_�c", j + 1
        Headder_Dic.Add "DPI_��", j + 2
        j = j + 2
    End If
        
    '�w�b�_�[���Z���Ɋi�[
    For Each c In Headder_Dic.Keys
        Cells(1, Headder_Dic(c)).Value = c
    Next
    
    '�i���o�[�N���X�̒�`
    Set objpb = New ProgressBar
    
    '�i���o�[�̃^�C�g����ݒ�
    objpb.SetTitle "�摜�̃X�e�[�^�X�`�F�b�N"
    
    
    '�e�t�@�C�����J���Ă���
    For Each Image_File In Image_Files
        '�Y���Z�����󔒂̏ꍇ�͉������Ȃ�
        If Image_File <> "" Then
            
            '�t�@�C���p�X
            Cells(i + 1, Headder_Dic("�t�@�C���p�X")).Value = Image_File
            
            '�y�[�W��
            If Page_flg Then Cells(i + 1, Headder_Dic("�y�[�W��")).Value = UBound(Split(img.Identify("-format", "%p,", Image_File), ","))
            
            '���k�`��
            If Compress_flg Then Cells(i + 1, Headder_Dic("���k�`��")).Value = Split(img.Identify("-format", "%C,", Image_File), ",")(0)
            
            '�s�N�Z����
            If Pixel_flg Then
                Cells(i + 1, Headder_Dic("�s�N�Z����_�c")).Value = Split(img.Identify("-format", "%h,", Image_File), ",")(0)
                Cells(i + 1, Headder_Dic("�s�N�Z����_��")).Value = Split(img.Identify("-format", "%w,", Image_File), ",")(0)
            End If
            
            'DPI
            If DPI_flg Then
                Cells(i + 1, Headder_Dic("DPI_�c")).Value = Split(img.Identify("-format", "%y,", Image_File), ",")(0)
                Cells(i + 1, Headder_Dic("DPI_��")).Value = Split(img.Identify("-format", "%x,", Image_File), ",")(0)
            End If
            
        End If
        
        '�v���O���X�o�[�X�V
        objpb.SetTitle "�摜�̃X�e�[�^�X�`�F�b�N�@" & i & "�F" & (File_Count + 1)
        objpb.SetProgress i / (File_Count + 1)
        
        '�i���̃J�E���g�A�b�v
        i = i + 1
    Next
    
    Set img = Nothing
    
    '�v���O���X�o�[�̔j��
    Set objpb = Nothing
    
    '�w�b�_�[�����̕��ŗ񕝂���������
    Range("A1:G1").Columns.AutoFit
    
    '�X�e�[�^�X�o�[�����Z�b�g
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    MsgBox ("�������������܂���")
    
End Sub
