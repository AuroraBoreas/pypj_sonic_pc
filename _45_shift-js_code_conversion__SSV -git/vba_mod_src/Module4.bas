Attribute VB_Name = "Module4"
Sub henkan()

Dim Cx0, Cx1, Cy0, Cy1, Cu0, Cu1, Cv0, Cv1, uv, JND As Double

On Error GoTo �G���[������

Sheets("��u'v'��JND�ϊ�").Range("c8:d8").Value = ""

Cx0 = Sheets("��u'v'��JND�ϊ�").Range("e8").Value '�Cx
Cy0 = Sheets("��u'v'��JND�ϊ�").Range("f8").Value '�Cy
Cx1 = Sheets("��u'v'��JND�ϊ�").Range("g8").Value '�Ώ�Cx
Cy1 = Sheets("��u'v'��JND�ϊ�").Range("h8").Value '�Ώ�Cy


Cu0 = 4 * Cx0 / (-2 * Cx0 + 12 * Cy0 + 3) '�Cu'
Cv0 = 9 * Cy0 / (-2 * Cx0 + 12 * Cy0 + 3) '�Cv'
Cu1 = 4 * Cx1 / (-2 * Cx1 + 12 * Cy1 + 3) '�Ώ�Cu'
Cv1 = 9 * Cy1 / (-2 * Cx1 + 12 * Cy1 + 3) '�Ώ�Cv'


''''��u'v'�v�Z''''

uv = ((Cu1 - Cu0) * (Cu1 - Cu0) + (Cv1 - Cv0) * (Cv1 - Cv0)) ^ (1 / 2) '��u'v'�l
Sheets("��u'v'��JND�ϊ�").Range("c8").Value = uv '��u'v'�l���ʕ\��


''''JND�v�Z''''

Sheets("JND�v�Z").Range("i16").Value = Cx0 'JND�v�Z�V�[�g���(�Cx)
Sheets("JND�v�Z").Range("i17").Value = Cx1 'JND�v�Z�V�[�g���(�Ώ�Cx)
Sheets("JND�v�Z").Range("j16").Value = Cy0 'JND�v�Z�V�[�g���(�Cy)
Sheets("JND�v�Z").Range("j17").Value = Cy1 'JND�v�Z�V�[�g���(�Ώ�Cy)

JND = Sheets("JND�v�Z").Range("k17").Value 'JND�l
Sheets("��u'v'��JND�ϊ�").Range("d8").Value = JND 'JND�l���ʕ\��

Exit Sub


''''�G���[����''''

�G���[������:

Sheets("��u'v'��JND�ϊ�").Range("c8:h8").Value = ""



End Sub
