Attribute VB_Name = "Module1"
'Attribute VB_Name = "ALG_Module6_1"
'***************************************************************************************************
'
'   Author:     by shinta
'***************************************************************************************************

Option Explicit

Dim pi      As Double
Dim dw_n_PI()  As Double      ' ��]�pTheta ��=0,��,2��,3��,4��,��� ����]����Index No.�� dw_n_PI(j)�Ɋi�[

Dim dw_deg  As Double           ' [deg] Division width for Wrap angle�i>0�j
Dim dw      As Double           ' Division width for Wrap angle�i>0�j
Dim dw_t_deg   As Double        ' Division number for Wrap thickness angle [deg] �i>0�j
Dim dw_t  As Double             ' Division number for Wrap thickness angle [rad] �i>0�j

Dim dw_c    As Long             ' Matrix number of colume�i>0�j
Dim dw_n    As Long             ' Matrix Division number�i>0�j
Dim dw_n_end As Long            ' Matrix number of wrap end�i>0�j
Dim div_n   As Long             ' Division number for Chamber Area�i>0�j
Dim Index_I As Long
Dim Matrix_A() As Double         ' equation Matorix A
Dim Matrix_C() As Double         ' equation Matorix C


