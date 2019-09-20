Attribute VB_Name = "Module1"
'Attribute VB_Name = "ALG_Module6_1"
'***************************************************************************************************
'
'   Author:     by shinta
'***************************************************************************************************

Option Explicit

Dim pi      As Double
Dim dw_n_PI()  As Double      ' 回転角Theta θ=0,π,2π,3π,4π,･･･ 半回転毎のIndex No.を dw_n_PI(j)に格納

Dim dw_deg  As Double           ' [deg] Division width for Wrap angle（>0）
Dim dw      As Double           ' Division width for Wrap angle（>0）
Dim dw_t_deg   As Double        ' Division number for Wrap thickness angle [deg] （>0）
Dim dw_t  As Double             ' Division number for Wrap thickness angle [rad] （>0）

Dim dw_c    As Long             ' Matrix number of colume（>0）
Dim dw_n    As Long             ' Matrix Division number（>0）
Dim dw_n_end As Long            ' Matrix number of wrap end（>0）
Dim div_n   As Long             ' Division number for Chamber Area（>0）
Dim Index_I As Long
Dim Matrix_A() As Double         ' equation Matorix A
Dim Matrix_C() As Double         ' equation Matorix C


