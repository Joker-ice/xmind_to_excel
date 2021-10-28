# 文件路径
# file_name = r'D:/pycharm-workspace/xmind_to_excel/test1.xmind'

# 是否生成p0用例
isp0 = False
# 是否转换excel
to_excel = True

# excel列名
title_name = [u"用例目录", u"用例名称", u"用例类型", u"优先级", u"用例描述", u"前提条件", u"步骤", u"期待结果"]
# excel 每列的填充值
col_value = [0, [1, '-', 2, '-', -3, '-', -2, '-', 'random.randint(0, 9999)'], '功能逻辑', -1, '', '', ['step'], -2]

# p0用例文件列名
p0_title_name = [u"编号", u"模块", u"分类", u"用例描述", u"预期结果", u"自测结果（pass/fail/未测）", u"负责人", u"备注"]
# p0用例文件 每列的填充值
p0_col_value = [['index'], 0, 1, ['step'], -2, '', '', '']
