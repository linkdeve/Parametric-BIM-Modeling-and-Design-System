"""
FreeCAD智能荷载插件v2.0
1.通过自然语言施加荷载，以三维箭头表示
2.为结构设置（固定端/铰接）支座
3.滑动调节荷载符号的比例大小
4.选择指定sap2000模型文件，将荷载同步布置到sap2000
"""

# ---------- 模块初始化 ----------
import sys
import os
from PySide2 import QtWidgets, QtCore
import re
import math
import requests
import json
import time
import threading
from collections import OrderedDict
from pathlib import Path
import win32com.client
from win32com.client import Dispatch, constants

sys.path.append(r"D:\miniconda3\Library\bin")
sys.path.append(r"D:\miniconda3\Scripts")
sys.path.append(r"D:\miniconda3\lib\site-packages")

# ---------- 配置文件 ----------
API_KEY = "sk-dd2060358bb945b9b263a40fef8965bb" 
MAX_RETRY = 3  # API调用重试次数


# ---------- 路径配置 ----------

# 获取FreeCAD安装目录

PLUGIN_DIR = os.path.dirname(__file__)

# 构建图标路径
ICON_FILE = os.path.normpath(r"D:/1files/what/SmartLoad/Icons/smart_load.png")

# 图标文件验证
if not os.path.exists(ICON_FILE):
    FreeCAD.Console.PrintError(f"图标文件未找到：{ICON_FILE}\n")
else:
    FreeCAD.Console.PrintLog(f"成功加载图标：{ICON_FILE}\n")

# --------- 支座类型定义 ---------
SUPPORT_TYPES = {
    "铰支座": {
        "color": (0.8, 0.2, 0.2),  # 红色
        "geometry": "sphere"        # 使用球体表示旋转自由度
    },
    "固定支座": {
        "color": (0.2, 0.2, 0.8),  # 蓝色
        "geometry": "cube"          # 使用立方体表示刚性固定
    }
}
# --------- 荷载类型定义 ---------
LOAD_TYPE_MAP = {
    '恒': ('DeadLoad', 1),  # 1对应SAP2000中的Dead荷载类型
    '活': ('LiveLoad', 2)   # 2对应Live荷载类型
    }

# ---------- 全局比例管理器 ----------
class ArrowScaleManager:
    """箭头比例管理单例"""
    _instance = None
    
    def __new__(cls):
        if not cls._instance:
            cls._instance = super().__new__(cls)
            cls._instance.scale_factor = 1.0
            cls._instance.arrow_objects = {}
        return cls._instance
    
    def register_arrow(self, arrow_obj, load_value, direction, node_pos):
        """注册箭头对象"""
        key = arrow_obj.Name
        self._instance.arrow_objects[key] = {
            'load_value': load_value,
            'direction': direction,
            'node_pos': node_pos,
            'timestamp': time.time()
        }
    
    def update_all(self, new_scale):
        """更新所有箭头比例"""
        self._instance.scale_factor = new_scale
        doc = FreeCAD.ActiveDocument
        
        for arrow_name in list(self._instance.arrow_objects.keys()):
            if not doc.getObject(arrow_name):
                del self._instance.arrow_objects[arrow_name]
                continue
                
            self._update_single(arrow_name)
        
        doc.recompute()
    
    def _update_single(self, arrow_name):
        """更新单个箭头"""
        arrow_info = self._instance.arrow_objects[arrow_name]
        arrow_obj = FreeCAD.ActiveDocument.getObject(arrow_name)
        
        # 计算新参数
        new_length = arrow_info['load_value'] * self._instance.scale_factor * 50
        direction = arrow_info['direction']
        node_pos = arrow_info['node_pos']
        
        # 更新杆部
        shaft = arrow_obj.Links[0]
        shaft.Radius = 5 * self._instance.scale_factor
        shaft.Height = new_length
        shaft.Placement = self._create_placement(
            node_pos, 
            direction,
            offset=-(1.2*new_length+50)
        )
        
        # 更新头部
        head = arrow_obj.Links[1]
        head.Radius1 = 15 * self._instance.scale_factor
        head.Height = 0.2 * new_length
        head.Placement = self._create_placement(
            node_pos,
            direction,
            offset=-(0.2*new_length +50)
        )
    
    @staticmethod
    def _create_placement(base_pos, direction, offset=0):
        """生成三维变换"""
        return FreeCAD.Placement(
            base_pos + FreeCAD.Vector(*direction) * offset,
            FreeCAD.Rotation(FreeCAD.Vector(0,0,1), FreeCAD.Vector(*direction))
        )

# ---------- 节点智能识别模块 ----------
class TopChordDetector:
    """上弦节点智能识别器"""
    @classmethod
    def detect(cls, doc, height_ratio=0.7, angle_tol=15):
        """
        主检测方法
        :param doc: FreeCAD文档对象
        :param height_ratio: 高度筛选阈值比例（0.6-0.9）
        :return: 排序后的上弦节点列表
        """
        # 获取所有顶点对象（根据FreeCAD对象类型过滤）
        nodes = [
            obj for obj in doc.Objects 
            if hasattr(obj, 'Shape') 
            and len(obj.Shape.Vertexes) == 1  # 仅含单个顶点
            and not obj.isDerivedFrom("Part::Compound")  # 排除复合对象
    ]
        
        if len(nodes) < 3:
            return nodes  # 节点过少时直接返回
            
        # 动态计算高度阈值
        z_coords = [n.Shape.Vertexes[0].Point.z for n in nodes]
        z_min, z_max = min(z_coords), max(z_coords)
        threshold_z = z_min + (z_max - z_min) * height_ratio
        
        # 筛选候选节点
        candidates = [n for n in nodes if n.Shape.Vertexes[0].Point.z >= threshold_z]
        
        # 直接去重并排序
        seen = OrderedDict()
        for node in candidates:
            if node.Name not in seen:
                seen[node.Name] = node
        return sorted(seen.values(), key=lambda n: n.Shape.Vertexes[0].Point.x)
    
class EndNodeDetector:
    """下弦端点智能识别器"""
    
    @classmethod
    def detect_ends(cls, doc, upper_chord_ratio=0.7, lower_chord_ratio=0.3):
        """
        :param upper_chord_ratio: 上弦节点高度比例阈值
        :param lower_chord_ratio: 下弦节点高度比例阈值
        :return: (left_node, right_node)
        """
        # 获取所有有效节点
        all_nodes = [
            obj for obj in doc.Objects 
            if hasattr(obj, 'Shape') 
            and len(obj.Shape.Vertexes) == 1
            and not obj.isDerivedFrom("Part::Compound")
        ]
        
        # 获取上弦节点
        upper_nodes = TopChordDetector.detect(doc, height_ratio=upper_chord_ratio)
        upper_positions = {tuple(n.Shape.Vertexes[0].Point) for n in upper_nodes}
        
        # 初步筛选下弦候选
        candidate_nodes = [
            n for n in all_nodes 
            if tuple(n.Shape.Vertexes[0].Point) not in upper_positions
        ]
        
        # 动态高度过滤
        lower_nodes = cls._filter_by_height(candidate_nodes, lower_chord_ratio)
        
        # 直接取排序后的两端点
        if len(lower_nodes) < 2:
            return None, None
        
        sorted_nodes = sorted(lower_nodes, 
                            key=lambda n: n.Shape.Vertexes[0].Point.x)
        return sorted_nodes[0], sorted_nodes[-1]

    @classmethod
    def _filter_by_height(cls, nodes, ratio):
        """基于相对高度的节点过滤"""
        if not nodes:
            return []
            
        z_values = [n.Shape.Vertexes[0].Point.z for n in nodes]
        z_min, z_max = min(z_values), max(z_values)
        threshold = z_min + (z_max - z_min) * ratio
        return [n for n in nodes if n.Shape.Vertexes[0].Point.z <= threshold]

# --------- 支座管理器类 ---------
class SupportManager:
    @classmethod
    def create_support(cls, node, support_type):
        """根据类型创建支座几何体"""
        original_pos = node.Shape.Vertexes[0].Point
        pos = original_pos - FreeCAD.Vector(0, 0, 50)
        if support_type == "铰支座":
            return cls._create_hinge(pos)
        elif support_type == "固定支座":
            return cls._create_fixed(pos)
        else:
            raise ValueError("未知支座类型")

    @staticmethod
    def _create_hinge(position):
        """创建铰支座几何体"""
        # 基础
        base = FreeCAD.ActiveDocument.addObject("Part::Box", "HingeBase")
        base.Length = 200
        base.Width = 200
        base.Height = 50
        base.Placement = FreeCAD.Placement(
            position - FreeCAD.Vector(100, 100, 50),
            FreeCAD.Rotation()
        )
        
        # 旋转球体
        pivot = FreeCAD.ActiveDocument.addObject("Part::Sphere", "HingePivot")
        pivot.Radius = 50
        pivot.Placement = FreeCAD.Placement(
            position + FreeCAD.Vector(0, 0, 20),
            FreeCAD.Rotation()
        )
        
        # 组合部件
        support = FreeCAD.ActiveDocument.addObject("Part::Compound", "HingeSupport")
        support.Links = [base, pivot]
        support.ViewObject.DiffuseColor = [SUPPORT_TYPES["铰支座"]["color"]] * len(support.Links)
        return support

    @staticmethod
    def _create_fixed(position):
        """创建固定支座几何体"""
        # 基础
        base = FreeCAD.ActiveDocument.addObject("Part::Cylinder", "FixedBase")
        base.Radius = 80
        base.Height = 50
        base.Placement = FreeCAD.Placement(
            position - FreeCAD.Vector(0, 0, 50),
            FreeCAD.Rotation()
        )
        
        # 固定块
        block = FreeCAD.ActiveDocument.addObject("Part::Box", "FixedBlock")
        block.Length = 80
        block.Width = 80
        block.Height = 40
        block.Placement = FreeCAD.Placement(
            position - FreeCAD.Vector(40, 40, 10),
            FreeCAD.Rotation()
        )
        
        # 组合部件
        support = FreeCAD.ActiveDocument.addObject("Part::Compound", "FixedSupport")
        support.Links = [base, block]
        support.ViewObject.DiffuseColor = [SUPPORT_TYPES["固定支座"]["color"]] * len(support.Links)
        return support

# ---------- 自然语言处理模块 ----------
class LoadInstructionParser:
    """荷载指令解析引擎（集成DeepSeek API）"""
    
    API_ENDPOINT = "https://api.deepseek.com/v1/chat/completions"
    DIRECTION_MAP = {
        "竖直向下": (0, 0, -1),
        "竖直向上": (0, 0, 1),
        "水平向右": (1, 0, 0),
        "水平向左": (-1, 0, 0),
        "横向": (0, 1, 0),
        "纵向": (1, 0, 0)
    }
    
    def __init__(self, api_key):
        self.api_key = api_key
    
    def parse(self, text):
        """
        带重试机制的解析方法
        :param text: 用户输入的指令文本
        :return: 结构化荷载数据（字典）或None
        """
        for attempt in range(MAX_RETRY):
            try:
                response = self._call_api(text)
                return self._format_response(response)
            except requests.exceptions.RequestException as e:
                if attempt == MAX_RETRY - 1:
                    FreeCAD.Console.PrintError(f"API请求失败: {str(e)}")
                    return None
                time.sleep(1)
    
    def _call_api(self, text):
        """安全调用API方法"""
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json"
        }
        
        system_prompt = """作为结构工程专家，请从以下指令提取荷载参数，按JSON格式返回：
        {
            "load_type": "恒/活",
            "position": {"type": "all/indexes/range/condition", "value": [...]},
            "magnitude": 数值（kN）,
            "direction": "方向描述"
        }"""
        
        return requests.post(
            self.API_ENDPOINT,
            headers=headers,
            json={
                "model": "deepseek-chat",
                "messages": [
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": text}
                ],
                "response_format": {"type": "json_object"},
                "temperature": 0.1
            },
            timeout=100  # 增加超时控制
        )
    
    def _format_response(self, response):
        """响应数据格式化和验证"""
        data = json.loads(response.json()['choices'][0]['message']['content'])
        
        # 数据完整性校验
        required_fields = ['load_type', 'position', 'magnitude', 'direction']
        if not all(field in data for field in required_fields):
            raise ValueError("API响应缺少必要字段")
        
        # 荷载类型验证
        if data['load_type'] not in ['恒', '活']:
            raise ValueError(f"无效荷载类型: {data['load_type']}")
        
        # 数值转换和范围检查
        try:
            data['magnitude'] = float(data['magnitude'])
            if data['magnitude'] <= 0:
                raise ValueError("荷载数值必须大于0")
        except (TypeError, ValueError):
            raise ValueError("无效荷载数值")
        
        # 方向向量转换
        direction = data['direction']
        if direction not in self.DIRECTION_MAP:
            raise ValueError(f"未知方向描述: {direction}")
        data['direction_vector'] = self.DIRECTION_MAP[direction]
        
        return data


def find_matching_node(sap_model, fc_point, tolerance=0.001):
    """智能坐标匹配节点"""
    try:
        # 正确获取节点列表
        ret, node_count, node_names = sap_model.PointObj.GetNameList()
        if ret != 0:
            raise Exception(f"获取节点列表失败，错误码：{ret}")
        
        # 遍历所有节点获取坐标
        node_coords = []
        for name in node_names:
            x, y, z = 0.0, 0.0, 0.0
            ret, x, y, z = sap_model.PointObj.GetCoordCartesian(name, x, y, z)
            if ret == 0:
                node_coords.append((name, x, y, z))
            else:
                FreeCAD.Console.PrintWarning(f"节点{name}坐标获取失败\n")

        # 精确坐标匹配
        sap_coord = (fc_point.x, fc_point.y, fc_point.z)
        for name, nx, ny, nz in node_coords:
            if (abs(nx - sap_coord[0]) < tolerance and
                abs(ny - sap_coord[1]) < tolerance and
                abs(nz - sap_coord[2]) < tolerance):
                FreeCAD.Console.PrintLog(f"匹配成功：{name} ({nx:.3f}, {ny:.3f}, {nz:.3f})\n")
                return name

        FreeCAD.Console.PrintWarning(
            f"未找到匹配节点 | 目标坐标：{sap_coord} | 容差：{tolerance}\n"
            f"最近节点：{node_coords[0][0]} ({node_coords[0][1]:.3f}, {node_coords[0][2]:.3f}, {node_coords[0][3]:.3f})\n"
        )
        return None

    except Exception as e:
        FreeCAD.Console.PrintError(f"节点匹配错误：{str(e)}\n")
        return None
    

# ---------- 荷载施加核心模块 ----------
class LoadApplicator:
    """荷载施加与可视化引擎"""

    def __init__(self, doc, scale_factor=50.0):
        self.doc = doc
        self.scale_factor = scale_factor  # 可视化比例 1kN=50mm
        self.base_radius = 5
        self.COLOR_SCHEME = {
    '恒': (1.0, 0.0, 0.0),  # 红色
    '活': (0.0, 0.0, 1.0)   # 蓝色
    }

    def apply(self, load_data):
        """主施加方法"""
        nodes = self._select_nodes(load_data['position'])
        if not nodes:
            raise ValueError("未找到符合条件的节点")
        
        for node in nodes:
            self._create_arrow(node, load_data)
            self.doc.recompute()
            self._sync_to_sap2000(node, load_data)

    def _sync_to_sap2000(self, node, load_data):
        """与SAP2000同步荷载"""
        try:
            # 获取连接
            sap_conn = SAP2000Connector()
            if not sap_conn._validate_connection():  # 验证连接是否有效
                sap_conn.reconnect()
            model = sap_conn.sap_model
            
            if not model:
                raise RuntimeError("SAP2000连接未建立")
            if model.GetModelIsLocked():
                raise RuntimeError("模型处于锁定状态，请先保存现有模型")

            # 验证荷载类型映射
            load_info = LOAD_TYPE_MAP.get(load_data['load_type'])
            if not load_info:
                raise ValueError(f"无效荷载类型: {load_data['load_type']}")
            load_name, load_type = load_info

            # 创建荷载模式
            ret, _, pattern_names = model.LoadPatterns.GetNameList()
            if ret != 0:
                raise Exception(f"获取荷载模式失败，错误码：{ret}")
                
            if load_name not in pattern_names:
                FreeCAD.Console.PrintLog(f"正在创建荷载模式：{load_name}\n")
                if model.LoadPatterns.Add(load_name, load_type, 0, True) != 0:
                    raise Exception("创建荷载模式失败")

            # 坐标匹配
            sap_node = find_matching_node(sap_conn.sap_model, node.Shape.Vertexes[0].Point)
            if not sap_node:
                raise ValueError(f"节点匹配失败 | FreeCAD坐标：{node.Shape.Vertexes[0].Point}")

            # 转换荷载向量
            direction = load_data['direction_vector']
            if not isinstance(direction, (list, tuple)) or len(direction) != 3:
                raise TypeError("方向向量格式错误，应为三元组")
                
            force = [load_data['magnitude'] * x for x in direction]

            ret = sap_conn.sap_model.PointObj.SetLoadForce(
                sap_node,       # 节点名称
                load_name,      # 荷载模式
                [force[0], force[1], force[2], 0.0, 0.0, 0.0],  # 6个元素的数组
                True,           # 覆盖现有荷载
                "Global",       # 坐标系
                0               # 相对位置
            )

            # 自动保存
            current_file = sap_conn.sap_model.GetModelFilename()
            save_path = str(current_file)
            ret = sap_conn.sap_model.File.Save(save_path)
            if ret != 0:
                raise Exception(f"自动保存失败，错误码：{ret}")
            FreeCAD.Console.PrintWarning(f"模型已保存至：{save_path}\n")

            FreeCAD.Console.PrintMessage("荷载同步成功！\n")

        except Exception as e:
            error_msg = f"SAP2000同步失败：{str(e)}"
            FreeCAD.Console.PrintError(error_msg + "\n")
            if 'ret' in locals():
                FreeCAD.Console.PrintError(f"详细错误码：{ret}\n")
            raise RuntimeError(error_msg) from e

    def _load_pattern_exists(self, sap_model, pattern_name):
        """荷载模式检查"""
        ret, num_patterns, pattern_names = sap_model.LoadPatterns.GetNameList()
        if ret != 0:
            raise Exception(f"获取荷载模式失败，错误码：{ret}")
        return pattern_name in pattern_names  # 正确访问名称列表
    
    def _select_nodes(self, position_info):
        """节点选择路由"""
        all_nodes = TopChordDetector.detect(self.doc)
        
        selector_type = position_info['type']

        if selector_type == 'all':
            return all_nodes
        elif selector_type == 'indexes':
            return self._select_by_index(all_nodes, position_info['value'])
        elif selector_type == 'condition':
            return self._filter_by_condition(all_nodes, position_info['value'])
        else:
            return []
    
    def _select_by_index(self, nodes, indexes):
        """索引选择器（1-based转0-based）"""
        valid_indexes = [i-1 for i in indexes if isinstance(i, int) and i > 0]
        return [nodes[i] for i in valid_indexes if i < len(nodes)]
    
    def _filter_by_condition(self, nodes, condition):
        """安全条件筛选器"""
        filtered = []
        for node in nodes:
            point = node.Shape.Vertexes[0].Point
            try:
                # 使用限制环境执行eval
                if eval(condition, {'X': point.x, 'Y': point.y, 'Z': point.z}, {}):
                    filtered.append(node)
            except:
                continue
        return filtered
    
    def _create_arrow(self, node, load_data):
        """参数化创建荷载箭头"""
        pos = node.Shape.Vertexes[0].Point
        direction = load_data['direction_vector']
        length = load_data['magnitude'] * self.scale_factor      #1kN对应50mm
        
        # 创建箭头组件
        shaft = self._create_shaft(pos, direction, length)
        head = self._create_head(pos, direction, length)
        
        # 组合对象并设置颜色
        if str(load_data['load_type']) == '恒':
            type = 'dead'
        else:
            type = 'live'
        load_name = type + str(load_data['magnitude'])
        compound = self.doc.addObject("Part::Compound", load_name)
        compound.Links = [shaft, head]
        compound.ViewObject.ShapeColor = self.COLOR_SCHEME[load_data['load_type']]
        del shaft, head

        # 注册到管理器
        ArrowScaleManager().register_arrow(
            compound,
            load_data['magnitude'],
            direction,
            pos
        )
        return compound
    
    def _create_shaft(self, pos, direction, length):
        """创建箭头杆部"""
        shaft = self.doc.addObject("Part::Cylinder", "LoadShaft")
        shaft.Radius = self.base_radius * self.scale_factor
        shaft.Height = length
        shaft.Placement = self._create_placement(pos, direction, offset=-(1.2*length+50))
        return shaft
    
    def _create_head(self, pos, direction, length):
        """创建箭头头部"""
        head = self.doc.addObject("Part::Cone", "LoadHead")
        head.Radius1 = self.base_radius * 3 * self.scale_factor
        head.Radius2 = 0
        head.Height = length * 0.2
        head.Placement = self._create_placement(
            pos, direction,  offset = -(0.2*length+50)
        )
        return head
    
    @staticmethod
    def _create_placement(base_pos, direction, offset=0, rotation_axis=FreeCAD.Vector(0,0,1)):
        """创建三维空间变换"""
        return FreeCAD.Placement(
            base_pos + FreeCAD.Vector(*direction) * offset,
             FreeCAD.Rotation(rotation_axis, FreeCAD.Vector(*direction))
        )
    
# ---------- 荷载加载对话界面 ----------
class LoadDialog(QtWidgets.QDialog):
    """荷载指令输入对话框"""
    
    def __init__(self, doc):
        """
        :param doc: FreeCAD文档对象（必须有效）
        """
        super().__init__()
        self.doc = doc
        self.parser = LoadInstructionParser(API_KEY)
        self._init_ui()
        self._connect_signals()
    
    def _init_ui(self):
        """界面初始化"""
        self.setWindowTitle("智能荷载工具")
        self.setMinimumSize(500, 200)
        layout = QtWidgets.QVBoxLayout()
        
        # 输入组件
        self.input = QtWidgets.QLineEdit()
        self.input.setPlaceholderText("示例：在前5个节点施加20kN水平向右的恒荷载")
        
        # 控制按钮
        self.btn_apply = QtWidgets.QPushButton("施加荷载")
        
        # 状态显示
        self.status = QtWidgets.QLabel()
        self.status.setAlignment(QtCore.Qt.AlignCenter)

        
        # 布局
        layout.addWidget(QtWidgets.QLabel("输入荷载指令："))
        layout.addWidget(self.input)
        layout.addWidget(self.btn_apply)
        layout.addWidget(self.status)
        self.setLayout(layout)
        
    
    def _connect_signals(self):
        """信号连接"""
        self.btn_apply.clicked.connect(self._on_apply)
    
    def _on_apply(self):
        """处理应用按钮点击"""
        text = self.input.text().strip()
        if not text:
            self._update_status("请输入有效指令", "warning")
            return
        
        try:
            load_data = self.parser.parse(text)
            if not load_data:
                raise ValueError("指令解析失败")
            
            applicator = LoadApplicator(self.doc)
            applicator.apply(load_data)
            self._update_status("荷载施加成功！", "success")
            FreeCADGui.activeDocument().activeView().viewIsometric()
        except Exception as e:
            self._update_status(f"错误：{str(e)}", "error")
    
    def _update_status(self, message, msg_type="info"):
        """更新状态信息"""
        color_map = {
            "info": "black",
            "success": "green",
            "warning": "orange",
            "error": "red"
        }
        self.status.setText(message)
        self.status.setStyleSheet(f"color: {color_map[msg_type]};")
    print('by link')

# --------- 支座选择对话框 ---------
class SupportDialog(QtWidgets.QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("支座类型选择")
        self.setFixedSize(300, 200)
        
        layout = QtWidgets.QVBoxLayout()
        
        # 左端选择
        left_layout = QtWidgets.QHBoxLayout()
        left_layout.addWidget(QtWidgets.QLabel("左端支座:"))
        self.left_combo = QtWidgets.QComboBox()
        self.left_combo.addItems(SUPPORT_TYPES.keys())
        left_layout.addWidget(self.left_combo)
        
        # 右端选择
        right_layout = QtWidgets.QHBoxLayout()
        right_layout.addWidget(QtWidgets.QLabel("右端支座:"))
        self.right_combo = QtWidgets.QComboBox()
        self.right_combo.addItems(SUPPORT_TYPES.keys())
        right_layout.addWidget(self.right_combo)
        
        # 操作按钮
        btn_box = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel
        )
        btn_box.accepted.connect(self.accept)
        btn_box.rejected.connect(self.reject)
        
        layout.addLayout(left_layout)
        layout.addLayout(right_layout)
        layout.addWidget(btn_box)
        self.setLayout(layout)

    def get_selections(self):
        """获取用户选择"""
        return {
            "left": self.left_combo.currentText(),
            "right": self.right_combo.currentText()
        }

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
            if current_units != 5:  # 5 = kN-mm
                self._sap_model.SetPresentUnits(5)
                FreeCAD.Console.PrintWarning("单位制已自动转换为kN-mm\n")

            # 验证模型状态
            if self._sap_model.GetModelIsLocked():
                raise ConnectionAbortedError("模型已被锁定（请检查SAP2000界面）")

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

# ---------- 添加命令 ----------

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

class AddSupportCommand:
    def GetResources(self):
        return {
            'Pixmap': '',
            'MenuText': "添加支座",
            'ToolTip': "为结构两端添加支座约束"
        }

    def Activated(self):
        if not FreeCAD.ActiveDocument:
            QtWidgets.QMessageBox.critical(None, "错误", "请先打开或创建文档")
            return
        
        # 检测端点
        try:
            left_node, right_node = EndNodeDetector.detect_ends(
            doc = FreeCAD.ActiveDocument,
            upper_chord_ratio=0.65,  # 确保完全排除上弦节点
            lower_chord_ratio=0.25    # 精确捕捉下弦端点
        )
        except Exception as e:
            QtWidgets.QMessageBox.critical(None, "检测异常", str(e))
            return

        # 添加空值检查
        if not (left_node and right_node):
            QtWidgets.QMessageBox.critical(None, "错误", "下弦端点识别失败")
            return

        # 显示选择对话框
        dialog = SupportDialog()
        if dialog.exec_() == QtWidgets.QDialog.Accepted:
            selections = dialog.get_selections()
            
            # 创建支座
            try:
                if selections["left"] != "无":
                    SupportManager.create_support(left_node, selections["left"])
                if selections["right"] != "无":
                    SupportManager.create_support(right_node, selections["right"])
                    
                FreeCAD.ActiveDocument.recompute()
                QtWidgets.QMessageBox.information(None, "成功", "支座添加完成")
            except Exception as e:
                QtWidgets.QMessageBox.critical(None, "错误", f"支座创建失败：{str(e)}")

    def IsActive(self):
        return FreeCAD.ActiveDocument is not None

class ApplyLoadCommand:
    """FreeCAD命令类"""

    def GetResources(self):
        return {
            'Pixmap': '',
            'MenuText': "智能荷载",
            'ToolTip': "启动荷载施加对话框",
        }

    def Activated(self):
        """命令激活时的处理"""
        if not FreeCAD.ActiveDocument:
            self._show_error("请先创建或打开文档")
            return
        
        dialog = LoadDialog(FreeCAD.ActiveDocument)
        dialog.exec_()
    
    def IsActive(self):
        """控制命令可用状态"""
        return FreeCAD.ActiveDocument is not None
    
    def _show_error(self, message):
        """显示错误消息"""
        QtWidgets.QMessageBox.critical(
            None,
            "操作错误",
            message,
            QtWidgets.QMessageBox.Ok
        )

class RealTimeScaleCommand:
    """实时比例调节控件"""
    
    def GetResources(self):
        return {
            'Pixmap': '',
            'MenuText': "实时比例",
            'ToolTip': "拖拽滑块实时调节箭头比例"
        }
    
    def Activated(self):
        # 创建浮动调节面板
        self.panel = QtWidgets.QWidget()
        self.panel.setWindowTitle("箭头比例调节")
        layout = QtWidgets.QVBoxLayout()
        
        # 添加滑动条
        self.slider = QtWidgets.QSlider(QtCore.Qt.Horizontal)
        self.slider.setRange(10, 500)  # 对应0.1到5.0倍
        self.slider.setValue(int(ArrowScaleManager().scale_factor * 50))
        self.slider.valueChanged.connect(self.update_scale)
        
        # 添加数值显示
        self.label = QtWidgets.QLabel()
        self.update_label()
        
        layout.addWidget(self.slider)
        layout.addWidget(self.label)
        self.panel.setLayout(layout)
        self.panel.show()
    
    def update_scale(self, value):
        # 更新比例值
        new_scale = value / 100.0
        ArrowScaleManager().update_all(new_scale)
        self.update_label()
    
    def update_label(self):
        self.label.setText(f"当前比例：{ArrowScaleManager().scale_factor:.1f}倍")

FreeCADGui.addCommand('ApplyLoad', ApplyLoadCommand())
FreeCADGui.addCommand('RealTimeScale', RealTimeScaleCommand())
FreeCADGui.addCommand('AddSupports', AddSupportCommand())
FreeCADGui.addCommand('SAPModelSelect', SAPModelSelector())

# ---------- FreeCAD工作台集成 ----------

class SmartLoadWorkbench(FreeCADGui.Workbench):
    """主工作台类"""

    MenuText = "智能荷载工具"
    ToolTip = "基于自然语言的荷载施加系统"
    Icon = ICON_FILE
    
    def Initialize(self):        
        # 创建工具栏和菜单
        self.appendToolbar("荷载工具", ["ApplyLoad",  "AddSupports", "SAPModelSelect", "RealTimeScale"])
        self.appendMenu("荷载工具", ["ApplyLoad", "AddSupports", "SAPModelSelect", "RealTimeScale"])
    
    def GetClassName(self):
        return "Gui::PythonWorkbench"


# ---------- 插件注册 ----------
FreeCADGui.addWorkbench(SmartLoadWorkbench())