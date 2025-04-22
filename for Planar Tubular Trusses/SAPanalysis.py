'''
桁架结构分析工具 -
功能：
1. 连接SAP2000模型
2. 双表格数据可视化

'''

import os
import sys
import FreeCAD
import FreeCADGui
from PySide import QtGui, QtCore
import math
import comtypes.client
import threading
import win32com.client
from dataclasses import dataclass
import json

sys.path.append(r"D:\miniconda3\Library\bin")
sys.path.append(r"D:\miniconda3\Scripts")
sys.path.append(r"D:\miniconda3\lib\site-packages")

# ---------------------------- 全局配置 ----------------------------
PLUGIN_DIR = os.path.dirname(__file__)
ICON_PATH = os.path.normpath(r"D:/1files/what/Analysis/Icons/sap_analysis.png")  # 标准化路径

# ---------- SAP200连接器 ----------
    
class SAP2000Connector:
    """SAP2000连接器"""
    _instance = None
    _is_connected = False
    _lock = threading.Lock()  # 线程安全锁

    def __new__(cls):
        with cls._lock:
            if not cls._instance:
                cls._instance = super().__new__(cls)
                # 初始化成员变量
                cls._instance._sap2000 = None
                cls._instance._sap_model = None
                cls._instance.model_path = None
                cls._instance._connection_attempts = 0
            return cls._instance

    @property
    def sap_model(self):
        return self._sap_model

    def set_model_path(self, path):
        """设置模型路径"""
        if not path:
            return False
            
        path = os.path.normpath(path)
        if not os.path.isfile(path):
            raise FileNotFoundError(f"模型文件不存在: {path}")
            
        if not path.lower().endswith(('.sdb', '.s2k')):
            raise ValueError("仅支持 .sdb 和 .s2k 格式")

        with self._lock:
            if self.model_path != path:
                self.model_path = path
                self._cleanup_connection()
                FreeCAD.Console.PrintLog(f"模型路径更新: {path}\n")
                
        return True

    def _validate_connection(self):
        """深度连接验证"""
        try:
            if not self._sap_model:
                return False
                
            # 通过获取基础信息验证连接
            ret, filename = self._sap_model.GetModelFilename()
            return ret == 0 and filename == self.model_path
        except:
            return False

    def _init_connection(self):
        """核心连接逻辑"""
        try:
            if self._is_connected:
                return True

            if not self.model_path:
                raise RuntimeError("未选择模型文件")

            self._sap2000 = win32com.client.Dispatch('CSI.SAP2000.API.SapObject')
            self._sap2000.ApplicationStart()
            self._sap_model = self._sap2000.SapModel

            # 打开模型文件
            ret = self._sap_model.File.OpenFile(self.model_path)
            if ret != 0:
                raise ConnectionError(f"打开模型失败 (错误码: {ret})")

            # 配置单位制
            current_units = self._sap_model.GetPresentUnits()
            if current_units != 9:  # 9 = N-mm
                self._sap_model.SetPresentUnits(9)
                FreeCAD.Console.PrintWarning("单位制已自动转换为N-mm\n")


            # 验证模型状态
        #   if self._sap_model.GetModelIsLocked():
        #        raise ConnectionAbortedError("模型已被锁定（请检查SAP2000界面）")

            self._is_connected = True
            self._connection_attempts = 0
            FreeCAD.Console.PrintMessage("成功连接SAP2000模型\n")
            return True
            
        except Exception as e:
            self._connection_attempts += 1
            self._cleanup_connection()
            if self._connection_attempts <= 3:
                FreeCAD.Console.PrintWarning(f"连接尝试失败 ({self._connection_attempts}/3): {str(e)}\n")
            else:
                raise RuntimeError("连接尝试超过最大次数") from e

    def select_model_interactive(self):
        """交互式文件选择"""
        from PySide2 import QtWidgets
        
        try:
            path, _ = QtWidgets.QFileDialog.getOpenFileName(
                None,
                "选择 SAP2000 模型文件",
                self._get_default_path(),
                "SAP2000 Files (*.sdb *.s2k);;All Files (*)"
            )
            if path and self.set_model_path(path):
                FreeCAD.Console.PrintMessage(f"已选择模型: {os.path.basename(path)}\n")
                return True
            return False
        except Exception as e:
            FreeCAD.Console.PrintError(f"文件选择失败: {str(e)}\n")
            return False

    def _get_default_path(self):
        """智能获取默认路径"""
        candidates = [
            os.path.join(os.environ.get("USERPROFILE", ""), "Desktop"),
            r"D:\DESKTOP\荷载施加模块",
            os.getcwd()
        ]
        for path in candidates:
            if os.path.exists(path):
                return path
        return ""

    def _cleanup_connection(self):
        """安全释放资源"""
        try:
            if self._sap_model:
                self._sap_model.File.Close()
        except Exception as e:
            FreeCAD.Console.PrintError(f"关闭模型文件失败: {str(e)}\n")
        try:
            if self._sap2000:
                self._sap2000.ApplicationExit(False)
        except Exception as e:
            FreeCAD.Console.PrintError(f"退出应用失败: {str(e)}\n")
        finally:
            self._sap2000 = None
            self._sap_model = None
            self._is_connected = False

    def reconnect(self):
        """带指数退避的重连机制"""
        import time
        max_retries = 3
        base_delay = 1  # 初始延迟1秒
        
        for attempt in range(max_retries):
            try:
                self._cleanup_connection()
                return self._init_connection()
            except Exception as e:
                if attempt == max_retries - 1:
                    raise
                sleep_time = base_delay * (2 ** attempt)
                FreeCAD.Console.PrintWarning(f"重连失败，{sleep_time}秒后重试...\n")
                time.sleep(sleep_time)
        return False

# ---------------------------- 数据加载器 ----------------------------
class TrussDataLoader(QtCore.QObject):
    """数据加载器"""

    finished = QtCore.Signal(dict)
    error = QtCore.Signal(str)

    def __init__(self):
        super().__init__()
        self.connector = SAP2000Connector()
        self._abort = False

    def run(self):
        """主加载流程"""
        try:
            if not self.connector._validate_connection():  # 验证连接是否有效
                self.connector.reconnect()
            model = self.connector.sap_model
            
            if not model:
                raise RuntimeError("SAP2000连接未建立")
        #    if model.GetModelIsLocked():
        #        raise RuntimeError("模型处于锁定状态，请先保存现有模型")
            
            data = {
                'nodes': self._load_nodes(),
                'frames': self._load_frames(),
                'forces': self._get_force() 
            }
            self.finished.emit(data)
        except Exception as e:
            self.error.emit(f"数据加载失败: {str(e)}")

    def _load_nodes(self):
        """加载节点数据"""
        nodes = {}
        ret, num_nodes, names = self.connector.sap_model.PointObj.GetNameList()
        
        for i, name in enumerate(names):
            if self._abort: 
                return {}
                
            # 获取坐标数据
            ret, x, y, z  = self.connector.sap_model.PointObj.GetCoordCartesian(name)

            nodes[name] = {
                'x': x,
                'y': y,
                'z': z,
                'frames': []
            }

        ret, num_frames, names = self.connector.sap_model.FrameObj.GetNameList()
        for i, name in enumerate(names):
            ret, p1, p2 = self.connector.sap_model.FrameObj.GetPoints(name)
            node1, node2 = str(p1), str(p2)
            nodes[node1]['frames'].append(name)
            nodes[node2]['frames'].append(name)

        return nodes

    def _load_frames(self):
        """加载杆件数据"""
        frames = {}
        ret, num_frames, names = self.connector.sap_model.FrameObj.GetNameList()
        nodes = self._load_nodes()  # 先加载节点数据
        
        for i, name in enumerate(names):
            if self._abort: 
                return {}
                
            # 获取杆件端点
            ret, p1, p2 = self.connector.sap_model.FrameObj.GetPoints(name)
            node1, node2 = str(p1), str(p2)

            # 计算长度
            dx = nodes[node2]['x'] - nodes[node1]['x']
            dy = nodes[node2]['y'] - nodes[node1]['y']
            dz = nodes[node2]['z'] - nodes[node1]['z']
            length = math.sqrt(dx**2 + dy**2 + dz**2)
            
            # 截面属性
            ret, prop, _ = self.connector.sap_model.FrameObj.GetSection(name)
            section = self._get_section(prop)
            # 计算角度
            angle = self._calc_angle(nodes[node1], nodes[node2])

            
            frames[name] = {
                'nodes': [node1, node2],
                'section': section,
                'length': length,
                'angle': angle
            }
            
        return frames

    def _get_section(self, prop_name):
        """获取截面属性"""
        mat_name = self.connector.sap_model.PropFrame.GetPipe(prop_name)[2]
        D = self.connector.sap_model.PropFrame.GetPipe(prop_name)[3]
        t = self.connector.sap_model.PropFrame.GetPipe(prop_name)[4]
        return {'material': mat_name, 'diameter': D, 'thickness': t, 'fy':235, 'fu':345}

    def _calc_angle(self, n1, n2):
        """计算空间角度"""
        dx = n2['x'] - n1['x']
        dy = n2['y'] - n1['y']
        dz = n2['z'] - n1['z']
        return math.degrees(math.atan2(dz, math.hypot(dx, dy)))
    print('by link')

    def _get_force(self):
        """获取杆件和节点的内力"""
        force_data = {
            'frame_forces': {},
            'joint_forces': {}
        }
        
        # 获取所有杆件名称
        ret, num_frames, frame_names = self.connector.sap_model.FrameObj.GetNameList()
        if ret != 0 or not frame_names:
            return force_data
        
        # 设置分析参数
        sap_model = self.connector.sap_model
        sap_model.SetPresentUnits(9)  # 设置为N-mm单位制
        sap_model.Results.Setup.DeselectAllCasesAndCombosForOutput()
        sap_model.Results.Setup.SetCaseSelectedForOutput("DeadLoad")
        
        # 获取杆件内力
        for frame_name in frame_names:
            try:
                # 杆件轴力 (拉正压负)
                raw_value = sap_model.Results.FrameForce(frame_name, 0)[9][0]
                force_data['frame_forces'][frame_name] = float(raw_value) if raw_value else 0.0
            except Exception as e:
                FreeCAD.Console.PrintError(f"杆件 {frame_name} 内力获取失败: {str(e)}\n")
                force_data['frame_forces'][frame_name] = 0.0
        
        # 获取节点内力
        for frame_name in frame_names:
            try:
                joints = sap_model.Results.FrameJointForce(frame_name, 0)
                for i, joint in enumerate(joints[4]):  # 遍历节点列表
                    key = f"{frame_name}||{joint}"
                    force_data['joint_forces'][key] = [
                        float(joints[8][i]) if joints[8][i] else 0.0,
                        float(joints[9][i]) if joints[9][i] else 0.0,
                        float(joints[10][i]) if joints[10][i] else 0.0
                    ]
            except Exception as e:
                FreeCAD.Console.PrintError(f"节点内力 {frame_name} 获取失败: {str(e)}\n")
        
        return force_data

# ---------------------------- 构件验算器 ----------------------------
class MemberCheck:
    """钢管构件验算"""

    # ----------------------------
    # 嵌套数据类定义
    # ----------------------------
    @dataclass
    class _HollowSection:
        """空心圆管截面参数"""
        D: float  # 外径(mm)
        t: float  # 壁厚(mm)

        def __post_init__(self):
            """数据有效性验证"""
            if self.t >= self.D / 2:
                raise ValueError("壁厚不能超过半径")
            if self.D <= 0 or self.t <= 0:
                raise ValueError("尺寸必须为正数")

    @dataclass
    class _SteelMaterial:
        """钢材材料属性"""
        fy: float  # 屈服强度(MPa)
        fu: float  # 抗拉强度(MPa)
        E: float = 206000  # 弹性模量 (MPa)，默认Q235~Q460钢

        def __post_init__(self):
            if self.fy <= 0 or self.fu <= 0:
                raise ValueError("强度参数必须为正数")

    @dataclass
    class _MemberForce:
        """构件内力"""
        N: float  # 轴力 (kN)，压力为负，拉力为正

        def __post_init__(self):
            if abs(self.N) > 1e6:
                raise ValueError("轴力值超出合理范围")

    @dataclass
    class _MemberGeometry:
        """构件几何参数"""
        L: float  # 几何长度(mm)
        μ: float = 1.0  # 计算长度系数（表7.4.1）

        def __post_init__(self):
            if self.L <= 0:
                raise ValueError("长度必须为正数")

    # ----------------------------
    # 初始化方法
    # ----------------------------
    def __init__(self,
                 D: float, t: float,
                 fy: float, fu: float,
                 N: float, L: float):
        """
        参数初始化入口[1](@ref)
        :param D: 外径(mm)
        :param t: 壁厚(mm)
        :param fy: 屈服强度(MPa)
        :param fu: 抗拉强度(MPa)
        :param N: 轴力(kN)
        :param L: 几何长度(mm)
        """
        # 实例化内部数据类
        self.section = self._HollowSection(D, t)
        self.material = self._SteelMaterial(fy, fu)
        self.force = self._MemberForce(N)
        self.geometry = self._MemberGeometry(L)

        # 计算截面特性
        self.section_props = self._calculate_section_properties()

        # 存储验算结果
        self.result = {
            'status': "通过",
            'checks': [],
            'N_Ed': abs(N)  # 轴力设计值（N）
        }

    # ----------------------------
    # 核心计算方法
    # ----------------------------
    def _calculate_section_properties(self) -> dict:
        """计算截面几何特性"""
        D, t = self.section.D, self.section.t
        d = D - 2 * t  # 内径
        A = math.pi * (D ** 2 - d ** 2) / 4  # 截面积 (mm²)
        I = math.pi * (D ** 4 - d ** 4) / 64  # 惯性矩 (mm⁴)
        i = math.sqrt(I / A)  # 回转半径
        return {'A': A, 'I': I, 'i': i}

    def _verify_strength(self):
        """执行强度验算"""
        σ = self.result['N_Ed'] / self.section_props['A']  # 压力强度设计值
        limit = self.material.fy if self.force.N < 0 else 0.7 * self.material.fu  # 拉力
        self._add_check('强度验算', σ, limit)

    def _verify_stability(self):
        """执行稳定性验算"""
        check_data = {
        'name': '稳定验算',
        'value': 'N/A',
        'limit': 'N/A',
        'passed': True,
        'note': '受拉构件不验算稳定性' if self.force.N >=0 else None
    }
        if self.force.N < 0:

            # 正则化长细比计算
            λ = (self.geometry.μ * self.geometry.L) / self.section_props['i']
            λn = (λ / math.pi) * math.sqrt(self.material.fy / self.material.E)

            # 稳定系数计算（b类截面参数）可根据后续需求查表更改
            α1, α2, α3 = 0.65, 0.965, 0.300  # 附录D表D.0.1
            φ = (1.0 - α1 * λn ** 2) if λn <= 0.215 else (
                    (α2 + α3 * λn + λn ** 2 - math.sqrt((α2 + α3 * λn + λn ** 2) ** 2 - 4 * λn ** 2)) / (2 * λn ** 2))

            σ_cr = φ * self.material.fy
            self._add_check('稳定验算',
                        self.result['N_Ed'] / self.section_props['A'],
                        σ_cr,
                        extra_data={'φ': φ, 'λn': λn})


    def _verify_slenderness(self):
        """长细比验算"""
        λ = (self.geometry.μ * self.geometry.L) / self.section_props['i']
        limit = 150 if self.force.N < 0 else 350
        self._add_check('长细比', λ, limit)

    def _add_check(self, name, value, limit, extra_data=None):
        """添加验算记录"""
        check = {
            'name': name,
            'value': value,
            'limit': limit,
            'passed': value <= limit
        }
        if extra_data:
            check.update(extra_data)
        self.result['checks'].append(check)

        # 更新整体状态
        if not check['passed']:
            self.result['status'] = "不通过"

    # ----------------------------
    # 公共接口方法
    # ----------------------------
    def perform_check(self) -> dict:
        """执行完整验算流程"""
        self._verify_strength()
        self._verify_stability()
        self._verify_slenderness()
        return self.result

    def print_results(self):
        """格式化输出结果"""
        print(f"验算状态: {self.result['status']}")
        print(f"轴力设计值: {abs(self.force.N)}N ({'压力' if self.force.N < 0 else '拉力'})")
        print("详细验算结果:")
        for check in self.result['checks']:
            print(f"\n{check['name']}")
            if 'MPa' in check['name']:
                print(f"  计算值: {check['value']:.1f} MPa")
                print(f"  允许值: {check['limit']:.1f} MPa")
            else:
                print(f"  计算值: {check['value']:.1f}")
                print(f"  允许值: {check['limit']}")

            if 'φ' in check:
                print(f"  稳定系数φ: {check['φ']:.3f}")
                print(f"  正则化长细比λn: {check['λn']:.3f}")
            print(f"  结论: {'满足' if check['passed'] else '不满足'}")

# ---------------------------- 基本数据框 ----------------------------
class DataDialog(QtGui.QDialog):
    """数据可视化对话框"""
    
    def __init__(self, data, parent=None):
        super().__init__(parent)
        self.data = data
        self._init_ui()
        
    def _init_ui(self):
        """初始化界面"""
        self.setWindowTitle("桁架数据查看器")
        self.setMinimumSize(1000, 600)
        
        layout = QtGui.QVBoxLayout()
        tabs = QtGui.QTabWidget()
        
        # 节点标签页
        node_tab = QtGui.QWidget()
        self._init_node_table(node_tab)
        tabs.addTab(node_tab, "节点信息")
        
        # 杆件标签页
        frame_tab = QtGui.QWidget()
        self._init_frame_table(frame_tab)
        tabs.addTab(frame_tab, "杆件信息")
        
        layout.addWidget(tabs)
        self.setLayout(layout)

    def _init_node_table(self, parent):
        """节点表格初始化"""
        table = QtGui.QTableWidget()
        table.setColumnCount(5)
        table.setHorizontalHeaderLabels(["节点", "X", "Y", "Z", "连接数"])
        table.verticalHeader().setVisible(False)
        table.setEditTriggers(QtGui.QAbstractItemView.NoEditTriggers)
        
        nodes = self.data['nodes']
        table.setRowCount(len(nodes))
        for row, (name, node) in enumerate(nodes.items()):
            table.setItem(row, 0, QtGui.QTableWidgetItem(name))
            table.setItem(row, 1, self._num_item(node['x']))
            table.setItem(row, 2, self._num_item(node['y']))
            table.setItem(row, 3, self._num_item(node['z']))
            table.setItem(row, 4, QtGui.QTableWidgetItem(str(len(node['frames']))))
        
        table.resizeColumnsToContents()
        parent.setLayout(QtGui.QVBoxLayout())
        parent.layout().addWidget(table)


    def _init_frame_table(self, parent):
        """杆件表格初始化"""
        table = QtGui.QTableWidget()
        table.setColumnCount(10)
        headers = ["杆件", "起点", "终点", "材料", "直径", "厚度", "角度","轴力(N)", "起点内力(Fx,Fy,Fz)(N)", "终点内力(Fx,Fy,Fz)(N)"]
        table.setHorizontalHeaderLabels(headers)
        table.verticalHeader().setVisible(False)
        table.setEditTriggers(QtGui.QAbstractItemView.NoEditTriggers)


        
        frames = self.data['frames']
        table.setRowCount(len(frames))
        
        for row, (name, frame) in enumerate(frames.items()):
            sect = frame.get('section', {})
            table.setItem(row, 0, QtGui.QTableWidgetItem(name))
            table.setItem(row, 1, QtGui.QTableWidgetItem(frame['nodes'][0]))
            table.setItem(row, 2, QtGui.QTableWidgetItem(frame['nodes'][1]))
            table.setItem(row, 3, QtGui.QTableWidgetItem(sect.get('material', 'N/A')))
            table.setItem(row, 4, self._num_item(sect.get('diameter', 0)))
            table.setItem(row, 5, self._num_item(sect.get('thickness', 0)))
            table.setItem(row, 6, self._num_item(frame.get('angle', 0), 1))

            # 轴力数据
            axial_force = self.data['forces']['frame_forces'].get(name, 'N/A')
            table.setItem(row, 7, self._num_item(axial_force, 0))
            
            node1, node2 = frame['nodes']

            # 节点力数据
            key_start = f"{name}||{node1}"
            key_end = f"{name}||{node2}"
            
            # 起点内力
            start_forces = self.data['forces']['joint_forces'].get(key_start, (0, 0, 0))
            table.setItem(row, 8, QtGui.QTableWidgetItem(
                f"({float(start_forces[0]):.0f}, {float(start_forces[1]):.0f}, {float(start_forces[2]):.0f})" 
            ))

            # 终点内力
            end_forces = self.data['forces']['joint_forces'].get(key_end, (0, 0, 0))
            table.setItem(row, 9, QtGui.QTableWidgetItem(
                f"({end_forces[0]:.0f}, {end_forces[1]:.0f}, {end_forces[2]:.0f})"
            ))
        
        table.resizeColumnsToContents()
        parent.setLayout(QtGui.QVBoxLayout())
        parent.layout().addWidget(table)

    def _num_item(self, value, prec=2):
        """数值格式化"""
        item = QtGui.QTableWidgetItem()
        item.setTextAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
        try:
            item.setText(f"{float(value):.{prec}f}")
        except:
            item.setText("N/A")
        return item

class ResultDialog(QtGui.QDialog):
    """显示验算结果的表格对话框"""
    
    def __init__(self, results, parent=None):
        super().__init__(parent)
        self.results = results
        self._init_ui()
        
    def _init_ui(self):

        # 安全获取检查项的方法
        def get_check(checks, name):
            for item in checks:
                if item['name'] == name:
                    return item
            return {'value': 'N/A', 'limit': 'N/A'}  # 默认值
        
        self.setWindowTitle("构件验算结果")
        self.setMinimumSize(1200, 800)
        
        layout = QtGui.QVBoxLayout()
        table = QtGui.QTableWidget()
        
        # 设置表格列
        headers = [
            "构件名称", "状态", "轴力(N)", 
            "强度验算(MPa)", "稳定验算(MPa)", "长细比"
        ]
        table.setColumnCount(len(headers))
        table.setHorizontalHeaderLabels(headers)
        
        # 填充数据
        table.setRowCount(len(self.results))

        for row, (name, result) in enumerate(self.results):
            table.setItem(row, 0, QtGui.QTableWidgetItem(name))
            table.setItem(row, 1, QtGui.QTableWidgetItem(result['status']))
            table.setItem(row, 2, QtGui.QTableWidgetItem(f"{abs(result['N_Ed']):.1f}"))
            
            # 提取各检查项结果
            strength = get_check(result['checks'], '强度验算')
            stability = get_check(result['checks'], '稳定验算')
            slenderness = get_check(result['checks'], '长细比')
            
            # 强度验算列
            strength_value = strength.get('value', 'N/A')
            strength_limit = strength.get('limit', 'N/A')
            strength_text = f"{strength_value:.2f}/{strength_limit:.2f}" if isinstance(strength_value, (float, int)) and isinstance(strength_limit, (float, int)) else f"{strength_value}/{strength_limit}"
            table.setItem(row, 3, QtGui.QTableWidgetItem(strength_text))
            
            # 稳定验算列
            stability_value = stability.get('value', 'N/A')
            stability_text = f"{stability_value:.2f}" if isinstance(stability_value, (float, int)) else f"{stability_value}"
            if 'φ' in stability:
                stability_text += f" (φ={stability['φ']:.3f})"
            elif stability.get('note'):
                stability_text += f" ({stability['note']})"
            table.setItem(row, 4, QtGui.QTableWidgetItem(stability_text))
            
            # 长细比列
            slenderness_value = slenderness.get('value', 'N/A')
            slenderness_limit = slenderness.get('limit', 'N/A')
            slenderness_text = f"{slenderness_value:.2f}/{slenderness_limit}" if isinstance(slenderness_value, (float, int)) else f"{slenderness_value}/{slenderness_limit}"
            table.setItem(row, 5, QtGui.QTableWidgetItem(slenderness_text))
        
        table.resizeColumnsToContents()
        layout.addWidget(table)
        self.setLayout(layout)

# ----------命令注册 ----------
class SAPModelSelector:
    """模型选择命令工具"""
    
    def GetResources(self):
        return {
            'Pixmap': '',
            'MenuText': "选择SAP模型",
            'ToolTip': "选择要连接的SAP2000模型文件"
        }

    def Activated(self):
        try:
            sap_conn = SAP2000Connector()
            if sap_conn.select_model_interactive():
                pass
        except Exception as e:
            FreeCAD.Console.PrintError(f"模型选择失败: {str(e)}\n")

    def IsActive(self):
        return True
    
class LoadDataCommand:
    """数据加载命令"""
    
    def GetResources(self):
        return {
            'Pixmap': '',
            'MenuText': "加载桁架数据",
            'ToolTip': "从SAP2000模型加载桁架数据",
        }

    def Activated(self):
        # 检查模型是否已加载
        connector = SAP2000Connector()
        if not connector.model_path:
            QtGui.QMessageBox.warning(None, "警告", "请先选择SAP2000模型文件")
            return
            
        # 创建加载线程
        self.loader = TrussDataLoader()
        self.thread = QtCore.QThread()
        
        # 连接信号
        self.loader.moveToThread(self.thread)
        self.loader.finished.connect(self._show_data)
        self.loader.error.connect(self._show_error)
        self.thread.started.connect(self.loader.run)

        self.thread.start()
        

    def _show_data(self, data):
        self.thread.quit()
        self.thread.wait()

        # 文档存储逻辑
        doc = FreeCAD.ActiveDocument
        if not hasattr(doc, 'TrussData'):
            # 使用PropertyString代替PropertyMap
            obj = doc.addObject("App::FeaturePython", "TrussData")
            obj.addProperty("App::PropertyString", "Data", "Internal", "桁架分析数据")
        else:
            obj = doc.TrussData
            
        # 序列化为JSON字符串
        serializable_data = json.dumps({
            'nodes': data['nodes'],
            'frames': data['frames'],
            'forces': data['forces']
        }, ensure_ascii=False)
        
        # 存储数据到文档
        obj.Data = serializable_data
        doc.recompute()

        DataDialog(data).exec_()

    def _show_error(self, message):
        self.thread.quit()
        self.thread.wait()  # 等待线程结束
        QtGui.QMessageBox.critical(None, "错误", message)

    def IsActive(self):
        return True

class CheckMembersCommand:
    """构件验算命令"""
    
    def GetResources(self):
        return {
            'Pixmap': '',
            'MenuText': "验算构件",
            'ToolTip': "执行构件强度验算"
        }

    def Activated(self):
        # 获取已加载的数据
        doc = FreeCAD.ActiveDocument
        if not hasattr(doc, 'TrussData'):
            QtGui.QMessageBox.critical(None, "错误", "请先加载模型数据")
            return
        
        # 从JSON字符串恢复数据
        try:
            raw_data = json.loads(doc.TrussData.Data)
            joint_forces = {}
            for key_str, value in raw_data['forces']['joint_forces'].items():
                frame_name, joint = key_str.split('||')
                joint_forces[(frame_name, joint)] = tuple(value)
           
        except Exception as e:
            QtGui.QMessageBox.critical(None, "数据错误", f"节点内力解析失败: {str(e)}")
            return
        
        data = {
                'nodes': raw_data['nodes'],
                'frames': raw_data['frames'],
                'forces': {
                    'frame_forces': raw_data['forces']['frame_forces'],
                    'joint_forces': joint_forces
                }
            }
        
        results = []
        
        # 遍历所有构件进行验算
        for frame_name, frame in data['frames'].items():
            try:
                # 获取内力
                N = data['forces']['frame_forces'].get(frame_name, 0)   # 单位为N
                
                checker = MemberCheck(
                    D=frame['section']['diameter'],
                    t=frame['section']['thickness'],
                    fy=frame['section']['fy'],
                    fu=frame['section']['fu'],
                    N=N,
                    L=frame['length']
                )
                result = checker.perform_check()
                results.append((frame_name, result))
                
            except Exception as e:
                FreeCAD.Console.PrintError(f"构件{frame_name}验算失败: {str(e)}\n")
        
        # 显示结果表格
        ResultDialog(results).exec_()

    def IsActive(self):
        return True

FreeCADGui.addCommand('CheckMembers', CheckMembersCommand())
FreeCADGui.addCommand('SAPModelSelect', SAPModelSelector())
FreeCADGui.addCommand('LoadData', LoadDataCommand())

# ---------------------------- FreeCAD集成 ----------------------------
class TrussWorkbench(FreeCADGui.Workbench):
    """FreeCAD工作台实现"""
    
    MenuText = "桁架验算"
    ToolTip = "桁架结构自动化验算工具"
    Icon = ICON_PATH if os.path.exists(ICON_PATH) else ""
    
    def Initialize(self):

        # 创建工具栏和菜单
        self.appendToolbar("桁架工具", ["SAPModelSelect", "LoadData", "CheckMembers"] )
        self.appendMenu("桁架分析", [ "SAPModelSelect", "LoadData", "CheckMembers"] )

    def GetClassName(self):
        return "Gui::PythonWorkbench"


# ---------------------------- 插件注册 ----------------------------

FreeCADGui.addWorkbench(TrussWorkbench())