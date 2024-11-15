# _*_ coding: utf-8 _*_
import rhinoscriptsyntax as rs
import scriptcontext as sc
import os

def GetSTLSettings(ang, ar, dist, grid, den, maxL=0, minL=0.0001):
    e_str = "_ExportFileAs=_Binary "
    e_str += "_ExportUnfinishedObjects=_Yes "
    e_str += "_UseSimpleDialog=_No "
    e_str += "_UseSimpleParameters=_{} ".format("Yes" if rs.ExeVersion() == 7 else "No")
    e_str += "_Enter _DetailedOptions "
    e_str += "_JaggedSeams=_No "
    e_str += "_PackTextures=_No "
    e_str += "_Refine=_Yes "
    e_str += "_SimplePlane=_Yes "
    e_str += "_AdvancedOptions "
    e_str += "_Angle={} ".format(ang)
    e_str += "_AspectRatio={} ".format(ar)
    e_str += "_Distance={} ".format(dist)
    e_str += "_Density={} ".format(den)
    e_str += "_Grid={} ".format(grid)
    e_str += "_MaxEdgeLength={} ".format(maxL)
    e_str += "_MinEdgeLength={} ".format(minL)
    e_str += "_Enter _Enter"
    return e_str

def BatchExportSTLByObject():
    us = rs.UnitSystem()
    # 2 - Millimeters (1.0e-3 meters)
    if us != 2:
        msg = "模型单位不是毫米"
        rs.MessageBox(msg)
        return

    msg = "选择需要导出为单独STL文件的几何体:"
    objs = rs.GetObjects(msg, 8 + 16 + 32, preselect=True)
    if not objs: return

    # 获取文件名，未保存为None
    doc_name = sc.doc.Name

    fileType = "STL Files (*.stl)|*.stl||"
    if not doc_name:
        rs.MessageBox("文件尚未保存,请保存文件后再执行程序")
        return
    else:
        # document has been saved, get path
        msg = "选择需要保存的目录："
        # Display browse-for-folder dialog allowing the user to select a folder
        folder = rs.BrowseForFolder(rs.WorkingFolder(), msg)
        if not folder: return
        filename = os.path.join(folder, doc_name)

    # 不同精度的转换参数
    # （最大角度，最大长宽比，边缘至曲面最大距离，网格面最小数目，密度，最大边缘长度，最小边缘长度）
    coarse_mm = [30, 0, 0.03, 16, 0, 0, 0.01]
    medium_mm = [15, 0, 0.01, 32, 0, 0, 0.001]
    fine_mm = [5, 0, 0.005, 64, 0, 0, 0.0001]
    extrafine_mm = [2, 0, 0.0025, 64, 0, 0, 0.0001]

    # 设置需要的精度类别，主要通过最小边长度控制
    a = '低精度' ; b = '中精度' ; c = '高精度' ; d = '超高精'
    if us == 2:
        exs = [a + ' (0.03mm)', b + ' (0.01mm)', c + ' (0.005mm)', d + ' (0.0025mm)']

    msg = "选择需要的精度"
    ex_choice = rs.ListBox(exs, msg, "STL export", exs[2])
    if ex_choice is None: return

    if ex_choice == exs[0]:
        sett_str = GetSTLSettings(*coarse_mm)
    elif ex_choice == exs[1]:
        sett_str = GetSTLSettings(*medium_mm)
    elif ex_choice == exs[2]:
        sett_str = GetSTLSettings(*fine_mm)
    elif ex_choice == exs[3]:
        sett_str = GetSTLSettings(*extrafine_mm)

    # 开始导出文件
    rs.EnableRedraw(False)
    for i, obj in enumerate(objs):
        objName = rs.ObjectName(obj)
        e_file_name = "{}-{}-{}.stl".format(filename[:-4], str(i+1), objName)
        print(e_file_name)
        rs.UnselectAllObjects()
        rs.SelectObject(obj)
        rs.Command('-_Export "{}" {} _Enter'.format(e_file_name, sett_str), True)
    rs.EnableRedraw(True)

BatchExportSTLByObject()