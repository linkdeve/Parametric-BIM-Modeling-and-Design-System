"""
平面管桁架参数化模型
1.在FC生成参数化的管桁架模型
2.在Sap2000同步生成结构模型，与FC实时同步
"""

# ---------- 库导入 ----------
import FreeCAD as App
import Part
import FreeCADGui as Gui
from FreeCAD import Base
from PySide2 import QtCore
import sys
import win32com.client
from win32com.client import Dispatch, constants

sys.path.extend([
    r"D:\miniconda3\Library\bin",
    r"D:\miniconda3\Scripts",
    r"D:\miniconda3\lib\site-packages"
])

# ---------- 参数化模型 ----------
class ParametricCircularTubeTruss:

    def __init__(self, obj):
        
        self.sap2000 = None
        self.sap_model = None
        self.is_sap_connected = False
        self.is_model_initialized = False
        obj.Proxy = self
        self.init_properties(obj)

    def init_properties(self, obj):
        # 几何属性
        obj.addProperty("App::PropertyLength", "Span", "Truss", "桁架跨度").Span = 5000.0
        obj.addProperty("App::PropertyLength", "Height", "Truss", "桁架高度").Height = 1000.0
        obj.addProperty("App::PropertyInteger", "Panels", "Truss", "节间数量").Panels = 5
        obj.addProperty("App::PropertyBool", "AlternateDiagonals", "Truss", "交替斜杆方向").AlternateDiagonals = True

        # 截面属性
        obj.addProperty("App::PropertyLength", "TopChordDiameter", "Top Chord", "上弦杆直径").TopChordDiameter = 50.0
        obj.addProperty("App::PropertyLength", "TopChordThickness", "Top Chord", "上弦杆壁厚").TopChordThickness = 5.0
        obj.addProperty("App::PropertyLength", "BottomChordDiameter", "Bottom Chord",
                        "下弦杆直径").BottomChordDiameter = 50.0
        obj.addProperty("App::PropertyLength", "BottomChordThickness", "Bottom Chord",
                        "下弦杆壁厚").BottomChordThickness = 5.0
        obj.addProperty("App::PropertyLength", "DiagonalDiameter", "Diagonals", "斜腹杆直径").DiagonalDiameter = 40.0
        obj.addProperty("App::PropertyLength", "DiagonalThickness", "Diagonals", "斜腹杆壁厚").DiagonalThickness = 4.0
        obj.addProperty("App::PropertyLength", "VerticalDiameter", "Verticals", "竖腹杆直径").VerticalDiameter = 30.0
        obj.addProperty("App::PropertyLength", "VerticalThickness", "Verticals", "竖腹杆壁厚").VerticalThickness = 3.0

        # 节点对象列表
        obj.addProperty("App::PropertyLinkList", "Nodes", "Truss", "节点对象列表")

    def onChanged(self, obj, prop):
        if prop in ["Span", "Height", "Panels", "TopChordDiameter", "TopChordThickness",
                    "BottomChordDiameter", "BottomChordThickness", "DiagonalDiameter",
                    "DiagonalThickness", "VerticalDiameter", "VerticalThickness", "AlternateDiagonals"]:
            self.execute(obj)
            self.update_sap2000(obj)

    def execute(self, obj):
        try:
            doc = obj.Document
            span = obj.Span.Value
            height = obj.Height.Value
            panels = obj.Panels

            # 生成节点坐标
            nodes = self.create_nodes(span, height, panels)

            # 清理旧节点
            if obj.Nodes:
                for node in obj.Nodes:
                    doc.removeObject(node.Name)
                obj.Nodes = []

            # 创建可见节点对象
            new_nodes = []
            for name, pos in nodes.items():
                node_obj = doc.addObject("Part::Vertex", "TrussNode")
                node_obj.Label = name  
                node_obj.X = pos.x
                node_obj.Y = pos.y
                node_obj.Z = pos.z
                node_obj.ViewObject.PointSize = 5
                node_obj.ViewObject.PointColor = (1.0, 0.0, 0.0) 
                new_nodes.append(node_obj)
            obj.Nodes = new_nodes

            # 创建杆件
            members = list(self.create_members(nodes, panels, obj.AlternateDiagonals, obj).values())
            obj.Shape = Part.Compound(members)

        except Exception as e:
            App.Console.PrintError(f"执行错误: {str(e)}\n")

    def create_nodes(self, span, height, panels):
        nodes = {}
        dx = span / panels
        # 下弦节点
        for i in range(panels + 1):
            nodes[f"L{i}"] = Base.Vector(i * dx, 0, 0)
        # 上弦节点
        for i in range(panels + 1):
            nodes[f"U{i}"] = Base.Vector(i * dx, 0, height)
        return nodes

    def create_members(self, nodes, panels, alternate, obj):
        members = {}
        # 上下弦杆
        for i in range(panels):
            # 下弦
            members[(f"L{i}", f"L{i + 1}")] = self.create_tube(
                nodes[f"L{i}"], nodes[f"L{i + 1}"],
                obj.BottomChordDiameter.Value, obj.BottomChordThickness.Value)
            # 上弦
            members[(f"U{i}", f"U{i + 1}")] = self.create_tube(
                nodes[f"U{i}"], nodes[f"U{i + 1}"],
                obj.TopChordDiameter.Value, obj.TopChordThickness.Value)
        # 竖腹杆
        for i in range(panels + 1):
            members[(f"L{i}", f"U{i}")] = self.create_tube(
                nodes[f"L{i}"], nodes[f"U{i}"],
                obj.VerticalDiameter.Value, obj.VerticalThickness.Value)
        # 斜腹杆
        for i in range(panels):
            if alternate and i % 2 == 1:
                members[(f"U{i}", f"L{i + 1}")] = self.create_tube(
                    nodes[f"U{i}"], nodes[f"L{i + 1}"],
                    obj.DiagonalDiameter.Value, obj.DiagonalThickness.Value)
            else:
                members[(f"L{i}", f"U{i + 1}")] = self.create_tube(
                    nodes[f"L{i}"], nodes[f"U{i + 1}"],
                    obj.DiagonalDiameter.Value, obj.DiagonalThickness.Value)
        return members

    def create_tube(self, p1, p2, diameter, thickness):
        direction = p2 - p1
        length = direction.Length
        direction.normalize()
        outer = Part.makeCylinder(diameter / 2, length, p1, direction)
        inner = Part.makeCylinder((diameter / 2) - thickness, length, p1, direction)
        return outer.cut(inner)

    def connect_to_sap2000(self):
        """连接或重新连接到 SAP2000。"""
        try:
            if not self.is_sap_connected:
                self.sap2000 = win32com.client.Dispatch('CSI.SAP2000.API.SapObject')
                self.sap2000.ApplicationStart()
                self.sap_model = self.sap2000.SapModel
                self.is_sap_connected = True
                App.Console.PrintMessage("SAP2000 连接已建立\n")
        except Exception as e:
            App.Console.PrintError(f"连接失败: {str(e)}\n")
            self.is_sap_connected = False

    def update_sap2000(self, obj):
        """将更新后的模型推送到 SAP2000。"""
        try:
            nodes = self.create_nodes(obj.Span.Value, obj.Height.Value, obj.Panels)
            members = self.create_members(nodes, obj.Panels, obj.AlternateDiagonals, obj)
            self.update_sap2000_model(obj, nodes, members)
        except Exception as e:
            App.Console.PrintError(f"更新失败: {str(e)}\n")

    def update_sap2000_model(self, obj, nodes, members):
        """调用 SAP2000 API 更新模型。"""
        try:
            self.connect_to_sap2000()
            if not self.is_sap_connected:
                raise ConnectionError("无法连接到 SAP2000")
            if not self.is_model_initialized:
                self.sap_model.InitializeNewModel()
                self.sap_model.SetPresentUnits(9)
                self.is_model_initialized = True
                self.sap_model.File.NewBlank()
            else:
                self.sap_model.File.NewBlank()

            for node_name, node_pos in nodes.items():
                self.sap_model.PointObj.AddCartesian(node_pos.x, node_pos.y, node_pos.z, node_name, node_name)

            for node, member in members.items():
                start_node = node[0]
                end_node = node[1]
                self.sap_model.FrameObj.AddByPoint(start_node, end_node, "1", "1", start_node+end_node)

            self.sap_model.PropMaterial.SetMaterial("Q235", 1)
            matName = "Q235"
            fy = 235
            fu = 345
            eFy = 235
            eFu = 345
            SSType = 1
            SSHysType = 2
            StrainAtHardening = 0.002
            StrainAtMaxStress = 0.02
            StrainAtRupture = 0.15
            FinalSlope = -0.1
            self.sap_model.PropMaterial.SetOSteel(matName, fy, fu, eFy, eFu, SSType, SSHysType,
                                      StrainAtHardening, StrainAtMaxStress, StrainAtRupture, FinalSlope)

            def set_pipe(name, d, t):
                if self.sap_model.PropFrame.SetPipe(name, "Q235", d, t) != 0:
                    App.Console.PrintError(f"截面 {name} 设置失败\n")

            set_pipe("TopChord", obj.TopChordDiameter.Value, obj.TopChordThickness.Value)
            set_pipe("BottomChord", obj.BottomChordDiameter.Value, obj.BottomChordThickness.Value)
            set_pipe("Diagonal", obj.DiagonalDiameter.Value, obj.DiagonalThickness.Value)
            set_pipe("Vertical", obj.VerticalDiameter.Value, obj.VerticalThickness.Value)

            for member_key in members.keys():
                start_node, end_node = member_key
                frame_name = start_node + end_node
                if start_node.startswith("U") and end_node.startswith("U"):
                    self.sap_model.FrameObj.SetSection(frame_name, "TopChord")
                elif start_node.startswith("L") and end_node.startswith("L"):
                    self.sap_model.FrameObj.SetSection(frame_name, "BottomChord")
                elif (start_node.startswith("L") and end_node.startswith("U")) and start_node[1:] != end_node[1:] or (
                        start_node.startswith("U") and end_node.startswith("L")) and start_node[1:] != end_node[1:] :
                    self.sap_model.FrameObj.SetSection(frame_name, "Diagonal")
                else:
                    self.sap_model.FrameObj.SetSection(frame_name, "Vertical")

                # 设置M2、M3端部释放（两端铰接）
                start_release = [False, False, False, False, True, True]  # U1,U2,U3,R1,R2,R3
                end_release = [False, False, False, False, True, True]
                start_value = [0.0] * 6  # 释放系数，全0表示完全释放
                end_value = [0.0] * 6
                ret = self.sap_model.FrameObj.SetReleases(frame_name, start_release, end_release, start_value, end_value, 0)
                    
                self.sap_model.Analyze.RunAnalysis()

        except Exception as e:
            App.Console.PrintError(f"SAP2000 操作失败: {str(e)}\n")
            self.is_sap_connected = False

    def __del__(self):
        if self.is_sap_connected:
            self.sap2000.ApplicationExit(False)
            App.Console.PrintMessage("SAP2000 连接已释放\n")

# ---------- 运行代码 ----------
def create_truss():
    """在控制台中创建桁架的快捷函数"""
    doc = App.ActiveDocument or App.newDocument()
    obj = doc.addObject("Part::FeaturePython", "Truss")
    ParametricCircularTubeTruss(obj)
    obj.ViewObject.Proxy = 0
    doc.recompute()
    Gui.SendMsgToActiveView("ViewFit")

create_truss()