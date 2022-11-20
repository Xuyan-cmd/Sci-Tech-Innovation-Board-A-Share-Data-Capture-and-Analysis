# coding:utf8
import time
import datetime
import threading
from io import StringIO
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfinterp import PDFResourceManager, process_pdf
from pathlib import Path
import pandas as pd


KEY_WORDS = ['大数据', '海量数据', '分布式计算', '数据集成', '元数据', '数据建模', '数据标准管理', '数据质量管理', '数据资产', '数据集市', '数据仓库', '数据标签', '主数据发现', '数据标准应用', '图分析', '图数据', '图结构数据', '图模型', '图计算', '知识图谱', 'BI工具', '商业智能', '数据可视化', '数据挖掘', '非结构化数据', '异构数据', 'Hadoop', 'Spark', '人工智能', 'AI', '机器学习', '深度学习', '强化学习', '迁移学习', '对抗学习', '监督学习', '多模态学习', '知识工程', '神经网络', '预训练模型', '自动驾驶', '数据标注', '语音识别', '图像识别', '机器翻译', '机器人', '计算机视觉', '自然语言处理', '智能推荐', '人脸识别', '图像识别', '智能化', '智能制造', '智能金融', '智能医疗', '智能安防', '智能交通',
             '智能医疗', '智慧城市', '智能农业', '物联网', 'IoT', '智能家居', '可穿戴设备', '传感器', 'eSIM', 'RFID', '电子标签', '边缘计算', '边缘终端', '边缘网关', '边缘控制器', '光伏云网', '车联网', '产业物联网', '工业互联网', '云计算', '云平台', '云服务', '公有云', '私有云', '混合云', '云端', 'IaaS', 'PaaS', 'SaaS', '云边协同', '云原生', '容器技术', '容器化', '微服务', 'DevOps', '虚拟现实', '增强现实', 'VR', 'AR', '人机交互', '感知交互', '近眼显示', '渲染计算', '云渲染', '注视点技术', '注视点光学', '注视点渲染', '眼动追踪', '手势追踪', '头显终端', '区块链', '联盟链', '公链', '公有链', '私有链', '数字货币', '比特币', 'BaaS', '密码算法', '对等式网络', '共识机制', '智能合约']

# 读取pdf


def read_pdf(pdf):
    # resource manager
    rsrcmgr = PDFResourceManager()
    retstr = StringIO()
    laparams = LAParams()
    # device
    device = TextConverter(rsrcmgr, retstr, laparams=laparams)
    process_pdf(rsrcmgr, device, pdf)
    device.close()
    content = retstr.getvalue()
    retstr.close()
    # 获取所有行
    # lines = str(content).split("\n")
    return content


if __name__ == '__main__':
    df = pd.read_excel('代码.xlsx')
    df = df.set_index('code', drop=True)
    p = Path().glob('*/*.pdf')

    year = 2021
    t_df = df.copy()
    for o in p:
        if str(year) in str(o):
            print(o)
        else:
            continue
        code = str(o).split('\\')[0]
        with open(o, 'rb') as o_pdf:
            text = read_pdf(o_pdf)
        key_words = t_df.columns[3:-1]
        for key_word in key_words:
            value = text.count(key_word)
            index_code = int(code)
            t_df.loc[index_code, key_word] = value
    t_df.to_excel('{}_cninfo_word_count.xlsx'.format(year))
