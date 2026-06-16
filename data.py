#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# @Time : 2020/8/26 14:48
# @Author : way
# @Site :
# @Describe:

import os
from collections import Counter, defaultdict
from functools import lru_cache

import xlrd


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_FILE = os.path.join(BASE_DIR, 'CulRelPro_China_1961-2019.xls')
_BATCH_ORDER = {
    '第一批': 1,
    '第二批': 2,
    '第三批': 3,
    '第四批': 4,
    '第五批': 5,
    '第六批': 6,
    '第七批': 7,
    '第八批': 8,
}
_PERIOD_GROUPS = [
    ('史前时期', ('旧石器时代', '新时期时代', '新石器时代')),
    ('先秦时期', ('夏', '商', '西周春秋', '战国')),
    ('秦汉时期', ('秦', '汉')),
    ('魏晋南北朝', ('三国', '晋', '南北朝')),
    ('隋唐五代', ('隋', '唐', '五代')),
    ('宋辽金元', ('宋', '辽', '金', '元')),
    ('明清时期', ('明', '清')),
    ('近现代', ('中华民国', '中华人民共和国')),
]
_RING_RADII = [
    ['59%', '70%'],
    ['49%', '60%'],
    ['39%', '50%'],
    ['29%', '40%'],
    ['20%', '30%'],
]
_MAP_ALIASES = {
    '北京市': '北京',
    '上海市': '上海',
    '天津市': '天津',
    '重庆市': '重庆',
    '香港特别行政区': '香港',
    '澳门特别行政区': '澳门',
    '大理白族自治州': '大理',
    '延边朝鲜族自治州': '延边',
    '湘西土家族苗族自治州': '湘西',
    '黔东南苗族侗族自治州': '黔东南',
    '黔西南布依族苗族自治州': '黔西南',
    '西双版纳傣族自治州': '西双版纳',
    '红河哈尼族彝族自治州': '红河',
    '德宏傣族景颇族自治州': '德宏',
    '楚雄彝族自治州': '楚雄',
    '怒江傈僳族自治州': '怒江',
    '文山壮族苗族自治州': '文山',
    '博尔塔拉蒙古自治州': '博尔塔拉',
    '巴音郭楞蒙古自治州': '巴音郭楞',
    '克孜勒苏柯尔克孜自治州': '克孜勒苏',
    '伊犁哈萨克自治州': '伊犁',
    '昌吉回族自治州': '昌吉',
}
_TYPE_DISPLAY_NAMES = {
    '近现代重要史迹及代表性建筑': '近现代史迹',
}


def _as_text(value):
    if value is None:
        return ''
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return str(value).strip()
    return str(value).strip()


def _as_int(value, default=0):
    if value in (None, ''):
        return default
    if isinstance(value, (int, float)):
        return int(value)
    try:
        return int(float(str(value).strip()))
    except (TypeError, ValueError):
        return default


def _load_sheet_rows(sheet, header_row=3, start_row=5, index_header='序号'):
    headers = [_as_text(sheet.cell_value(header_row, c)) for c in range(sheet.ncols)]
    rows = []
    index_col = None
    if index_header in headers:
        index_col = headers.index(index_header)
    for r in range(start_row, sheet.nrows):
        if index_col is not None:
            index_value = sheet.cell_value(r, index_col)
            if not isinstance(index_value, (int, float)):
                continue
        row = {}
        has_value = False
        for c, header in enumerate(headers):
            if not header:
                continue
            value = sheet.cell_value(r, c)
            if value not in ('', None):
                has_value = True
            row[header] = value
        if has_value:
            rows.append(row)
    return rows


def _display_name(name, mapping=None):
    name = _as_text(name)
    if not mapping:
        return name
    return mapping.get(name, name)


def _top_items(counter, limit=None, name_mapping=None):
    items = sorted(counter.items(), key=lambda item: (-item[1], item[0]))
    if limit is not None:
        items = items[:limit]
    return [
        {'name': _display_name(name, name_mapping), 'value': int(value)}
        for name, value in items if name and value
    ]


def _period_items(rows):
    period_counter = Counter()
    for row in rows:
        period_name = _as_text(row.get('时代（中文）'))
        period_value = _as_int(row.get('数量'))
        if period_name and period_value:
            period_counter[period_name] += period_value

    grouped_items = []
    used_periods = set()
    for group_name, aliases in _PERIOD_GROUPS:
        total = sum(period_counter.get(alias, 0) for alias in aliases)
        used_periods.update(aliases)
        if total:
            grouped_items.append({'name': group_name, 'value': total})

    remainder = sum(
        value for name, value in period_counter.items()
        if name not in used_periods
    )
    if remainder:
        grouped_items.append({'name': '其他时期', 'value': remainder})

    return grouped_items


def _batch_items(counter):
    items = sorted(
        counter.items(),
        key=lambda item: (_BATCH_ORDER.get(item[0], 999), item[0])
    )
    return [{'name': name, 'value': int(value)} for name, value in items if name and value]


def _normalize_map_name(name):
    name = _as_text(name)
    if not name:
        return ''
    if name in _MAP_ALIASES:
        return _MAP_ALIASES[name]
    for suffix in ('特别行政区', '自治州', '自治县', '自治旗', '地区', '盟', '市'):
        if name.endswith(suffix) and len(name) > len(suffix):
            normalized = name[:-len(suffix)]
            return _MAP_ALIASES.get(normalized, normalized)
    return name


@lru_cache(maxsize=4)
def _load_excel_dashboard(excel_mtime):
    workbook = xlrd.open_workbook(EXCEL_FILE)
    tab1_rows = _load_sheet_rows(workbook.sheet_by_name('Tab.1'))
    tab2_rows = _load_sheet_rows(workbook.sheet_by_name('Tab.2'))
    tab3_rows = _load_sheet_rows(workbook.sheet_by_name('Tab.3'))

    matched_geo_names = {
        '海门', '鄂尔多斯', '招远', '舟山', '齐齐哈尔', '盐城', '赤峰', '青岛', '乳山', '金昌',
        '泉州', '莱西', '日照', '胶南', '南通', '拉萨', '云浮', '梅州', '文登', '上海',
        '攀枝花', '威海', '承德', '厦门', '汕尾', '潮州', '丹东', '太仓', '曲靖', '烟台',
        '福州', '瓦房店', '即墨', '抚顺', '玉溪', '张家口', '阳泉', '莱州', '湖州', '汕头',
        '昆山', '宁波', '湛江', '揭阳', '荣成', '连云港', '葫芦岛', '常熟', '东莞', '河源',
        '淮安', '泰州', '南宁', '营口', '惠州', '江阴', '蓬莱', '韶关', '嘉峪关', '广州',
        '延安', '太原', '清远', '中山', '昆明', '寿光', '盘锦', '长治', '深圳', '珠海',
        '宿迁', '咸阳', '铜川', '平度', '佛山', '海口', '江门', '章丘', '肇庆', '大连',
        '临汾', '吴江', '石嘴山', '沈阳', '苏州', '茂名', '嘉兴', '长春', '胶州', '银川',
        '张家港', '三门峡', '锦州', '南昌', '柳州', '三亚', '自贡', '吉林', '阳江', '泸州',
        '西宁', '宜宾', '呼和浩特', '成都', '大同', '镇江', '桂林', '张家界', '宜兴', '北海',
        '西安', '金坛', '东营', '牡丹江', '遵义', '绍兴', '扬州', '常州', '潍坊', '重庆',
        '台州', '南京', '滨州', '贵阳', '无锡', '本溪', '克拉玛依', '渭南', '马鞍山', '宝鸡',
        '焦作', '句容', '北京', '徐州', '衡水', '包头', '绵阳', '乌鲁木齐', '枣庄', '杭州',
        '淄博', '鞍山', '溧阳', '库尔勒', '安阳', '开封', '济南', '德阳', '温州', '九江',
        '邯郸', '临安', '兰州', '沧州', '临沂', '南充', '天津', '富阳', '泰安', '诸暨',
        '郑州', '哈尔滨', '聊城', '芜湖', '唐山', '平顶山', '邢台', '德州', '济宁', '荆州',
        '宜昌', '义乌', '丽水', '洛阳', '秦皇岛', '株洲', '石家庄', '莱芜', '常德', '保定',
        '湘潭', '金华', '岳阳', '长沙', '衢州', '廊坊', '菏泽', '合肥', '武汉', '大庆'
    }

    province_counter = Counter()
    type_counter = Counter()
    batch_counter = Counter()
    city_counter = Counter()
    matched_map_counter = Counter()
    province_type_counter = defaultdict(Counter)

    for row in tab1_rows:
        type_name = _as_text(row.get('类型（中文）'))
        batch_name = _as_text(row.get('批次（中文）'))
        province_name = _as_text(row.get('省级政区名称（中文）'))
        city_name = _as_text(row.get('市级政区名称（中文）'))
        map_name = _normalize_map_name(city_name)

        if type_name:
            type_counter[type_name] += 1
        if batch_name:
            batch_counter[batch_name] += 1
        if province_name:
            province_counter[province_name] += 1
            if type_name:
                province_type_counter[province_name][type_name] += 1
        if city_name:
            city_counter[city_name] += 1
        if map_name and map_name in matched_geo_names:
            matched_map_counter[map_name] += 1

    total_units = len(tab1_rows)
    province_rows = [
        {
            'name': _as_text(row.get('省级政区名称（中文）')),
            'value': _as_int(row.get('数量')),
        }
        for row in tab2_rows
        if _as_text(row.get('省级政区名称（中文）')) and _as_int(row.get('数量')) > 0
    ]
    province_rows = sorted(province_rows, key=lambda item: (-item['value'], item['name']))

    period_rows = _period_items(tab3_rows)
    batch_rows = _batch_items(batch_counter)
    city_top_rows = _top_items(city_counter, limit=8)
    type_rows = _top_items(type_counter, limit=6, name_mapping=_TYPE_DISPLAY_NAMES)
    province_top_rows = province_rows[:10]
    province_compare_rows = province_rows[:8]

    compare_types = ['古建筑', '古遗址']
    echart4_data = {
        'title': '文保大省：古建筑 vs 古遗址',
        'xAxis': [item['name'] for item in province_compare_rows],
        'data': [
            {
                'name': type_name,
                'value': [
                    int(province_type_counter[item['name']].get(type_name, 0))
                    for item in province_compare_rows
                ]
            }
            for type_name in compare_types
        ]
    }

    ring_rows = batch_rows[:5]
    echart6_rows = []
    for idx, item in enumerate(ring_rows):
        value = int(item['value'])
        echart6_rows.append({
            'name': item['name'],
            'value': value,
            'value2': max(total_units - value, 0),
            'color': f'{idx + 1:02d}',
            'radius': _RING_RADII[idx],
        })

    map_rows = _top_items(matched_map_counter, limit=120)
    map_max = max([item['value'] for item in map_rows], default=100)
    map_symbol_size = max(int(map_max / 12), 8)

    return {
        'title': '全国重点文物保护单位数据大屏（1961-2019）',
        'counter': {
            'name': '全国重点文物保护单位总数',
            'value': total_units,
        },
        'counter2': {
            'name': '覆盖省级行政区数量',
            'value': len(province_rows),
        },
        'echart1_data': {
            'title': '国保单位类型分布',
            'data': type_rows,
        },
        'echart2_data': {
            'title': '国保单位省份分布',
            'data': province_top_rows,
        },
        'echarts3_1_data': {
            'title': '历史时期结构',
            'data': period_rows,
        },
        'echarts3_2_data': {
            'title': '公布批次分布',
            'data': batch_rows,
        },
        'echarts3_3_data': {
            'title': '国保单位集中城市TOP8',
            'data': city_top_rows,
        },
        'echart4_data': echart4_data,
        'echart5_data': {
            'title': '文保大省TOP8',
            'data': province_top_rows[:8],
        },
        'echart6_data': {
            'title': '公布批次占比',
            'data': echart6_rows,
        },
        'map_1_data': {
            'symbolSize': map_symbol_size,
            'data': map_rows,
        }
    }


class SourceDataDemo:

    def __init__(self):
        self.title = '大数据可视化展板通用模板'
        self.counter = {'name': '2018年总收入情况', 'value': 12581189}
        self.counter2 = {'name': '2018年总支出情况', 'value': 3912410}
        self.echart1_data = {
            'title': '行业分布',
            'data': [
                {"name": "商超门店", "value": 47},
                {"name": "教育培训", "value": 52},
                {"name": "房地产", "value": 90},
                {"name": "生活服务", "value": 84},
                {"name": "汽车销售", "value": 99},
                {"name": "旅游酒店", "value": 37},
                {"name": "五金建材", "value": 2},
            ]
        }
        self.echart2_data = {
            'title': '省份分布',
            'data': [
                {"name": "浙江", "value": 47},
                {"name": "上海", "value": 52},
                {"name": "江苏", "value": 90},
                {"name": "广东", "value": 84},
                {"name": "北京", "value": 99},
                {"name": "深圳", "value": 37},
                {"name": "安徽", "value": 150},
            ]
        }
        self.echarts3_1_data = {
            'title': '年龄分布',
            'data': [
                {"name": "0岁以下", "value": 47},
                {"name": "20-29岁", "value": 52},
                {"name": "30-39岁", "value": 90},
                {"name": "40-49岁", "value": 84},
                {"name": "50岁以上", "value": 99},
            ]
        }
        self.echarts3_2_data = {
            'title': '职业分布',
            'data': [
                {"name": "电子商务", "value": 10},
                {"name": "教育", "value": 20},
                {"name": "IT/互联网", "value": 20},
                {"name": "金融", "value": 30},
                {"name": "学生", "value": 40},
                {"name": "其他", "value": 50},
            ]
        }
        self.echarts3_3_data = {
            'title': '兴趣分布',
            'data': [
                {"name": "汽车", "value": 4},
                {"name": "旅游", "value": 5},
                {"name": "财经", "value": 9},
                {"name": "教育", "value": 8},
                {"name": "软件", "value": 9},
                {"name": "其他", "value": 9},
            ]
        }
        self.echart4_data = {
            'title': '时间趋势',
            'data': [
                {"name": "安卓", "value": [3, 4, 3, 4, 3, 4, 3, 6, 2, 4, 2, 4, 3, 4, 3, 4, 3, 4, 3, 6, 2, 4, 4]},
                {"name": "IOS", "value": [5, 3, 5, 6, 1, 5, 3, 5, 6, 4, 6, 4, 8, 3, 5, 6, 1, 5, 3, 7, 2, 5, 8]},
            ],
            'xAxis': ['01', '02', '03', '04', '05', '06', '07', '08', '09', '11', '12', '13', '14', '15', '16', '17',
                      '18', '19', '20', '21', '22', '23', '24'],
        }
        self.echart5_data = {
            'title': '省份TOP',
            'data': [
                {"name": "浙江", "value": 2},
                {"name": "上海", "value": 3},
                {"name": "江苏", "value": 3},
                {"name": "广东", "value": 9},
                {"name": "北京", "value": 15},
                {"name": "深圳", "value": 18},
                {"name": "安徽", "value": 20},
                {"name": "四川", "value": 13},
            ]
        }
        self.echart6_data = {
            'title': '一线城市情况',
            'data': [
                {"name": "浙江", "value": 80, "value2": 20, "color": "01", "radius": ['59%', '70%']},
                {"name": "上海", "value": 70, "value2": 30, "color": "02", "radius": ['49%', '60%']},
                {"name": "广东", "value": 65, "value2": 35, "color": "03", "radius": ['39%', '50%']},
                {"name": "北京", "value": 60, "value2": 40, "color": "04", "radius": ['29%', '40%']},
                {"name": "深圳", "value": 50, "value2": 50, "color": "05", "radius": ['20%', '30%']},
            ]
        }
        self.map_1_data = {
            'symbolSize': 100,
            'data': [
                {'name': '北京', 'value': 320},
                {'name': '上海', 'value': 320},
                {'name': '广州', 'value': 520},
                {'name': '深圳', 'value': 700},
                {'name': '成都', 'value': 380},
                {'name': '重庆', 'value': 350},
                {'name': '杭州', 'value': 420},
                {'name': '武汉', 'value': 360},
                {'name': '南京', 'value': 400},
                {'name': '天津', 'value': 320},
                {'name': '西安', 'value': 340},
                {'name': '郑州', 'value': 310},
                {'name': '长沙', 'value': 290},
                {'name': '青岛', 'value': 280},
                {'name': '沈阳', 'value': 270},
                {'name': '大连', 'value': 260},
                {'name': '济南', 'value': 250},
                {'name': '哈尔滨', 'value': 240},
                {'name': '福州', 'value': 230},
                {'name': '厦门', 'value': 320},
                {'name': '昆明', 'value': 210},
                {'name': '合肥', 'value': 200},
                {'name': '南昌', 'value': 195},
                {'name': '石家庄', 'value': 190},
                {'name': '太原', 'value': 185},
                {'name': '南宁', 'value': 180},
                {'name': '长春', 'value': 175},
                {'name': '温州', 'value': 170},
                {'name': '宁波', 'value': 165},
                {'name': '苏州', 'value': 160},
                {'name': '无锡', 'value': 155},
                {'name': '贵阳', 'value': 150},
                {'name': '珠海', 'value': 145},
                {'name': '兰州', 'value': 140},
                {'name': '洛阳', 'value': 135},
                {'name': '海口', 'value': 130},
                {'name': '乌鲁木齐', 'value': 320},
                {'name': '扬州', 'value': 120},
                {'name': '南通', 'value': 115},
                {'name': '烟台', 'value': 110},
                {'name': '海门', 'value': 105},
            ]
        }

    @property
    def echart1(self):
        data = self.echart1_data
        echart = {
            'title': data.get('title'),
            'xAxis': [i.get("name") for i in data.get('data')],
            'series': [i.get("value") for i in data.get('data')]
        }
        return echart

    @property
    def echart2(self):
        data = self.echart2_data
        echart = {
            'title': data.get('title'),
            'xAxis': [i.get("name") for i in data.get('data')],
            'series': [i.get("value") for i in data.get('data')]
        }
        return echart

    @property
    def echarts3_1(self):
        data = self.echarts3_1_data
        echart = {
            'title': data.get('title'),
            'xAxis': [i.get("name") for i in data.get('data')],
            'data': data.get('data'),
        }
        return echart

    @property
    def echarts3_2(self):
        data = self.echarts3_2_data
        echart = {
            'title': data.get('title'),
            'xAxis': [i.get("name") for i in data.get('data')],
            'data': data.get('data'),
        }
        return echart

    @property
    def echarts3_3(self):
        data = self.echarts3_3_data
        echart = {
            'title': data.get('title'),
            'xAxis': [i.get("name") for i in data.get('data')],
            'data': data.get('data'),
        }
        return echart

    @property
    def echart4(self):
        data = self.echart4_data
        echart = {
            'title': data.get('title'),
            'names': [i.get("name") for i in data.get('data')],
            'xAxis': data.get('xAxis'),
            'data': data.get('data'),
        }
        return echart

    @property
    def echart5(self):
        data = self.echart5_data
        echart = {
            'title': data.get('title'),
            'xAxis': [i.get("name") for i in data.get('data')],
            'series': [i.get("value") for i in data.get('data')],
            'data': data.get('data'),
        }
        return echart

    @property
    def echart6(self):
        data = self.echart6_data
        echart = {
            'title': data.get('title'),
            'xAxis': [i.get("name") for i in data.get('data')],
            'data': data.get('data'),
        }
        return echart

    @property
    def map_1(self):
        data = self.map_1_data
        sym = data.get('symbolSize')
        if sym is None or (isinstance(sym, (int, float)) and not (sym > 0)):
            sym = 100
        echart = {
            'symbolSize': sym,
            'data': data.get('data'),
        }
        return echart

    def to_dict(self):
        """
        将数据对象转换为字典格式，用于 JSON 序列化
        """
        return {
            'title': self.title,
            'counter': self.counter,
            'counter2': self.counter2,
            'echart1': self.echart1,
            'echart2': self.echart2,
            'echarts3_1': self.echarts3_1,
            'echarts3_2': self.echarts3_2,
            'echarts3_3': self.echarts3_3,
            'echart4': self.echart4,
            'echart5': self.echart5,
            'echart6': self.echart6,
            'map_1': self.map_1,
        }


class SourceData(SourceDataDemo):

    def __init__(self):
        """
        按照 SourceDataDemo 的格式覆盖数据即可
        """
        super().__init__()
        excel_mtime = os.path.getmtime(EXCEL_FILE)
        data = _load_excel_dashboard(excel_mtime)
        self.title = data.get('title')
        self.counter = data.get('counter')
        self.counter2 = data.get('counter2')
        self.echart1_data = data.get('echart1_data')
        self.echart2_data = data.get('echart2_data')
        self.echarts3_1_data = data.get('echarts3_1_data')
        self.echarts3_2_data = data.get('echarts3_2_data')
        self.echarts3_3_data = data.get('echarts3_3_data')
        self.echart4_data = data.get('echart4_data')
        self.echart5_data = data.get('echart5_data')
        self.echart6_data = data.get('echart6_data')
        self.map_1_data = data.get('map_1_data')
