import matplotlib.pyplot as plt, numpy as np, scipy.stats as stats, pydotplus, xlrd, xlwt
from sklearn import tree
from sklearn.cluster import KMeans
from sklearn.model_selection import train_test_split

plt.rcParams['font.sans-serif'] = ['SimHei']
plt.rcParams['axes.unicode_minus'] = False
addfiles = ['附件1： 123家有信贷记录企业的相关数据.xlsx',
            '附件2： 302家无信贷记录企业的相关数据.xlsx',
            '附件3： 银行贷款年利率与客户流失率关系的统计数据.xlsx'
    , '附件一数据分析结果.xls', '附件二数据分析结果.xls', '企业各年收入支出信息.xls',
            '附件1的离散化信息.xls', '附件2的离散化信息.xls']


# 读取文件
def loadxls(file, sheet=0, str=0, stc=0, enc=None, enr=0):
    re = []
    d = xlrd.open_workbook(addfiles[file - 1], 'rb')
    table = d.sheets()[sheet]  # 通过索引顺序获取
    nrows = table.nrows
    for i in range(str, nrows - enr):
        re.append(table.row_values(i, start_colx=stc, end_colx=enc))
    return re


# 保存文件
def savexls(d, file, cols):
    wb = xlwt.Workbook(encoding='utf-8')
    ws = wb.add_sheet('Sheet1')
    ws.write(0, 0, label=file)
    for i, t in enumerate(cols):
        ws.write(1, i, label=cols[i])
    for i, t in enumerate(d):
        for j, k in enumerate(d[0]):
            ws.write(i + 2, j, label=d[i][j])
    wb.save(file + '.xls')


# 依据税率和企业类型区分信贷政策
def brandclu(file):
    clus = [['3', '劳务', '建筑', '个体经营', '影', '管理', '五金', '设计服务', '维修'],  # 3%
            ['6', '房地产', '教育', '发展', '科学', '策划', '质量', '设计', '传播', '广告', '投资', '代理',
             '印务', '餐饮', '服务',
             '律师', '事务', '实业', '物资', '运', '物流', '汽贸', '快递', '环保', '地质', '灾害', '土地',
             '生态', '猕猴桃', '园艺',
             '园林'],  # 6%
            ['9', '图书', '石化', '天然气', '装饰', '农业', '调味品', '蔬菜', '建设'],  # 9%
            ['13', '建材', '包装', '材', '木业', '纸业', '塑胶', '经营', '童装', '纺织品', '生活用品', '鞋',
             '服饰', '体育', '家居',
             '卫浴', '门窗', '设备', '空调', '家电', '器材', '电器', '机电', '花', '轮胎', '化工', '食品'],  # 13%
            ['16', '电子', '技术', '机械', '工贸', '车', '工程', '网络', '文化', '贸易', '电气', '合金',
             '塑料', '商贸', '医疗', '药',
             '科技']]  # 16%
    for i, t in enumerate(file):
        flag = False
        for k, c in enumerate(clus):
            if flag:
                break
            for j in c:
                if j in t[1]:
                    if len(file[i]) == 2:
                        file[i].extend(['', ''])
                else:
                    file[i][2] = (ord(file[i][2]) - ord("A")) / 3
                if file[i][3] == '是':
                    file[i][3] = 1
                else:
                    file[i][3] = 0
                file[i].append(int(clus[k][0]))
                flag = True
                break
    return file


# 分析企业交易信息
def get_tradeinfo(compinf, buyfile, cellfile):
    labels = ["A", "B", "C", "D"]
    fracs = [27, 38, 34, 24]
    plt.pie(x=fracs, labels=labels, autopct="%0.2f%%")
    plt.title('附件一企业信誉评级分布')
    plt.show()


# [交易数 交易额 取消率]
def getinfo(file):
    inf = [[0, 0, 0]]
    comp = 0
    comptitle = file[0][0]
    for i in file:
        if i[0] != comptitle:
            inf[comp][2] = inf[comp][2] / inf[comp][0]
        comp += 1
        comptitle = i[0]
        inf.append([0, 0, 0])
        inf[comp][0] += 1
        inf[comp][1] += float(i[4])
        if i[7] == '作废发票':
            inf[comp][2] += 1
            inf[comp][2] = inf[comp][2] / inf[comp][0]
            return inf
    buyinf = getinfo(buyfile)
    cellinf = getinfo(cellfile)
    for i in range(len(buyinf)):
        compinf[i].extend(
            [cellinf[i][1] - buyinf[i][1], buyinf[i][1], cellinf[i][1], 100 * (buyinf[i][2] +
                                                                               cellinf[i][2]) / 2])
    return compinf


# 分析上下游影响度
def get_influ(compinf, buyfile, cellfile):
    guests = np.zeros((2, 42000), np.int16)
    bef = 0
    for i in buyfile:
        if i[7] == '有效发票' and (bef == 0 or bef != i[3]):
            bef = i[3]
        guests[0][int(i[3][1:])] += 1
        comp = 0
        comptitle = buyfile[0][0]
        buyinf = [0]
        brands = set()
        for i in buyfile:
            if i[7] != '有效发票':
                continue
        if i[0] != comptitle:
            for j in brands:
                buyinf[comp] += guests[0][j]
        buyinf[comp] = len(brands) / buyinf[comp] * 100
        comp += 1
        comptitle = i[0]
        buyinf.append(0)
        brands = set()
        brands.add(int(i[3][1:]))
        for j in brands:
            buyinf[comp] += guests[0][j]
        buyinf[comp] = len(brands) / buyinf[comp] * 100
        comp = 0
        comptitle = cellfile[0][0]
        cellinf = [0]
        guests = np.zeros((1, 42000), np.float)
        for i in cellfile:
            if i[7] != '有效发票':
                continue
        if i[0] != comptitle:
            t = 0
        for j in range(3):
            t += np.max(guests[0])
        d = np.where(guests[0] == np.max(guests[0]))
        guests[0][d[0][0]] = 0
        cellinf[comp] = t / cellinf[comp] * 100
        comp += 1
        comptitle = i[0]
        cellinf.append(0)
        guests = np.zeros((1, 42000), np.float)
        guests[0][int(i[3][1:])] += float(i[4])
        cellinf[comp] += float(i[4])
        t = 0
        for j in range(3):
            t += np.max(guests[0])
        d = np.where(guests[0] == np.max(guests[0]))
        guests[0][d[0][0]] = 0
        cellinf[comp] = t / cellinf[comp] * 100
        for i, t in enumerate(compinf):
            compinf[i].extend([buyinf[i], cellinf[i], (buyinf[i] + cellinf[i]) / 2])
    return compinf


# gm模型灰度预测各个企业的今年收益率
class GrayFore():

    def __init__(self, d):

    self.d = d


self.fore = d.copy()


def build_model(self):
    X_0 = np.array(self.fore)


X_1 = np.zeros(X_0.shape)
for i in range(X_0.shape[0]):
    X_1[i] = np.sum(X_0[0:i + 1])
Z_1 = np.zeros(X_1.shape[0] - 1)
for i in range(1, X_1.shape[0]):
    Z_1[i - 1] = -0.5 * (X_1[i] + X_1[i - 1])
B = np.append(np.array(np.mat(Z_1).T), np.ones(Z_1.shape).reshape((Z_1.shape[0], 1)), axis=1)
Yn = X_0[1:].reshape((X_0[1:].shape[0], 1))
B = np.mat(B)
Yn = np.mat(Yn)
a_ = (B.T * B) ** -1 * B.T * Yn
a, b = np.array(a_.T)[0]
X_ = np.zeros(X_0.shape[0])


def f(k):
    return (X_0[0] - b / a) * (1 - np.exp(a)) * np.exp(-a * (k))


self.fore.append(f(X_.shape[0]))


def forecast(self, time=1):
    for i in range(time):
        self.build_model()


return self.fore.copy()


# 问题一： 企业基本信息获取代码
def ques1_1():
    comp_inf1 = brandclu(loadxls(1, 0, 1))


comp_inf2 = brandclu(loadxls(2, 0, 1))
buyfile1 = loadxls(1, 1, 1)
cellfile1 = loadxls(1, 2, 1)
buyfile2 = loadxls(2, 2, 1)
cellfile2 = loadxls(2, 1, 1)
comp_inf1 = get_tradeinfo(comp_inf1, buyfile1, cellfile1)
comp_inf2 = get_tradeinfo(comp_inf2, buyfile2, cellfile2)
comp_inf1 = get_influ(comp_inf1, buyfile1, cellfile1)
comp_inf2 = get_influ(comp_inf2, buyfile2, cellfile2)
if 1:
    savexls(comp_inf1, '附件一数据分析结果',
            ['公司编号', '公司名称', '信用等级', '是否违约', '经营类型(税额)', '净收益', '总支出', '总收入',
             '交易取消率',
             '上游客户集中度', '下游客户集中度', '平均客户集中度'])
savexls(comp_inf2, '附件二数据分析结果',
        ['公司编号', '公司名称', '信用等级', '是否违约', '经营类型(税额)', '净收益', '总支出', '总收入',
         '交易取消率',
         '上游客户集中度', '下游客户集中度', '平均客户集中度'])


# 问题一： 企业还款额度
def ques1_2():


# 提取年份信息
comp_inf = [loadxls(4, 0, 2), loadxls(5, 0, 2)]
buyfile = [loadxls(1, 1, 1), loadxls(2, 2, 1)]
cellfile = [loadxls(1, 2, 1), loadxls(2, 1, 1)]
prof = []
for i in range(426):
    prof.append([0, 0, 0, 0, 0, 0])
for now_comp in range(2):
    for i, t in enumerate(buyfile[now_comp]):
        if t[7] != '有效发票':
        continue
year = None
if float(t[2]) >= 42736 and float(t[2]) < 43101:
    year = 0
elif float(t[2]) >= 43101 and float(t[2]) < 43466:
    year = 1
elif float(t[2]) >= 42466 and float(t[2]) < 43831:
year = 2
if year == None:
    continue
prof[int(t[0][1:])][year] += float(t[4])
for i, t in enumerate(cellfile[now_comp]):
    if t[7] != '有效发票':
        continue
year = None
if float(t[2]) >= 42736 and float(t[2]) < 43101:
    year = 0
elif float(t[2]) >= 43101 and float(t[2]) < 43466:
    year = 1
elif float(t[2]) >= 42466 and float(t[2]) < 43831:
year = 2
if year == None:
    continue
prof[int(t[0][1:])][3 + year] += float(t[4])
for i, t in enumerate(comp_inf[now_comp]):
    pro18 = prof[i + 1][4] - prof[i + 1][1]
pro19 = prof[i + 1][5] - prof[i + 1][2]
if pro18 <= 0:
    if pro19 <= 0:
        comp_inf[now_comp][i][-1] = (-1 - pro19 / pro18)
else:
    comp_inf[now_comp][i][-1] = (1 + pro19 / pro18)
else:
if pro19 <= 0:
    comp_inf[now_comp][i][-1] = (-pro19 / pro18)
else:
    comp_inf[now_comp][i][-1] = (pro19 / pro18)
if 1:
    savexls(comp_inf[0], '附件一数据分析结果',
            ['公司编号', '公司名称', '信用等级', '是否违约', '经营类型(税额)', '净收益', '总支出', '总收入',
             '交易取消率',
             '上游客户集中度', '下游客户集中度', '企业发展趋势'])
savexls(comp_inf[1], '附件二数据分析结果',
        ['公司编号', '公司名称', '信用等级', '是否违约', '经营类型(税额)', '净收益', '总支出', '总收入',
         '交易取消率',
         '上游客户集中度', '下游客户集中度', '企业发展趋势'])
savexls(prof[1:], '企业各年收入支出信息', ['17支出', '18支出', '19支出', '17收入', '18收入',
                                           '19收入'])


# 问题一： 企业情况聚类与风险评估确定
def ques1_3():


# 手肘法
comp_inf = [loadxls(4, 0, 2), loadxls(5, 0, 2)]
cho_col = [5, 6, 7, 8, 9, 10, 11]
best = [4, 5, 5, 4, 4, 4, 6,
        4, 4, 4, 4, 5, 4, 4]
col_name = ['净收益', '总支出', '总收入', '交易取消率', '上游客户集中度', '下游客户集中度',
            '企业发展趋势']
for now_comp in range(2):
    for col in cho_col:
        d = np.zeros((len(comp_inf[now_comp]), 1))
for i, t in enumerate(comp_inf[now_comp]):
    if int(comp_inf[now_comp][i][4]) == 3:
        comp_inf[now_comp][i][4] = 0
elif int(comp_inf[now_comp][i][4]) == 6:
comp_inf[now_comp][i][4] = 0.25
elif int(comp_inf[now_comp][i][4]) == 9:
comp_inf[now_comp][i][4] = 0.5
elif int(comp_inf[now_comp][i][4]) == 13:
comp_inf[now_comp][i][4] = 0.75
elif int(comp_inf[now_comp][i][4]) == 16:
comp_inf[now_comp][i][4] = 1
d[i][0] = t[col]
cons = [0, 0, 0, 0, 0, 0, 0, 0]
for i in range(2, 10):
    clf = KMeans(n_clusters=i)
clf = clf.fit(d)
cons[i - 2] = clf.inertia_
if best[now_comp * 7 + col - 5] != 0:
    clf = KMeans(n_clusters=best[now_comp * 7 + col - 5])
clf = clf.fit(d)
lab = clf.labels_
cen = clf.cluster_centers_
cen = [float(i) for i in cen]
t = cen.copy()
t.sort()
d = {}
for i, k1 in enumerate(t):
    for j, k2 in enumerate(cen):
        if k1 == k2:
        d[i] = j
break
for i, t in enumerate(lab):
    comp_inf[now_comp][i][col] = d[t] / (best[now_comp * 7 + col - 5] - 1)
else:
    plt.plot(list(range(2, 10)), cons)
plt.xlabel('k值')
plt.ylabel('质心距离平方和')
plt.title('附件' + str(now_comp + 1) + '企业' + col_name[col - 5] + 'K-Means聚类手肘图')
plt.show()
# 计算风险
wei = [0.29, 0, 0.06, -0.12, -0.12, -0.12, 0, 0.145, 0.145, -0.12]
mi = 1
ma = 0
for i in range(1):
    for j, t1 in enumerate(comp_inf[i]):
        sco = 0
for k, t2 in enumerate(t1[2:12]):
    sco += wei[k] * t2
if sco > ma:
    ma = sco
if sco < mi:
    mi = sco
comp_inf[i][j].append(sco)
for j, t1 in enumerate(comp_inf[i]):
    comp_inf[i][j].append(1 - (comp_inf[i][j][12] - mi) / (ma - mi))
savexls(comp_inf[now_comp], '附件' + str(now_comp + 1) + '的离散化信息',
        ['公司编号', '公司名称', '信用等级', '是否违约', '经营类型(税额)', '净收益', '总支出', '总收入',
         '交易取消率',
         '上游客户集中度', '下游客户集中度', '企业发展趋势', '信贷风险', '归一化风险'])


# 问题一： 企业最大贷款额度与贷款利率确定
def ques1_4():


# 20年现金流预测
prof = loadxls(6, 0, 2, 0, 6)
for i, t in enumerate(prof):
    t1 = [t[3] - t[0], t[4] - t[1], t[5] - t[2]]
t2 = [t[3], t[4], t[5]]
mod1 = GrayFore(t1)
mod2 = GrayFore(t2)
if i == 1:
    t = mod1.forecast()
plt.plot(list(range(2017, 2020)), t[:3])
plt.plot(list(range(2019, 2021)), t[2:])
plt.title('企业E' + str(i + 1) + '的2020年净收益灰度预测图')
plt.xlabel('年份')
plt.ylabel('收益额')
plt.xticks(list(range(2017, 2021)))
plt.legend(['历史数据', '预测数据'])
plt.show()
t = mod2.forecast()
plt.plot(list(range(2017, 2020)), t[:3])
plt.plot(list(range(2019, 2021)), t[2:])
plt.title('企业E' + str(i + 1) + '的2020年进项发票金额灰度预测图')
plt.xlabel('年份')
plt.ylabel('进项发票金额')
plt.legend(['历史数据', '预测数据'])
plt.xticks(list(range(2017, 2021)))
plt.show()
prof[i].append(mod1.forecast()[3] * 0.5)
prof[i].append(mod2.forecast()[3] * 0.05)
prof[i].append(max([prof[i][-1], prof[i][-2]]))
savexls(prof, '企业各年收入支出信息', ['17支出', '18支出', '19支出', '17收入', '18收入',
                                       '19收入', '预测收益', '预测进款额度', '最大额度'])
# 最大偿还力度预测
compinf = loadxls(7, 0, 2, 0, 14)
for i, t in enumerate(compinf):
    compinf[i].append(prof[i][-1])
compinf[i].append(float(prof[i][-1]) * float(compinf[i][-2]))
if compinf[i][-1] > 1000000:
    compinf[i][-1] = 1000000
elif compinf[i][-1] < 100000:
    compinf[i][-1] = 100000
# 风险聚类
clf = KMeans(n_clusters=3)
d = np.zeros((len(compinf), 1))
for i, t in enumerate(compinf):
    d[i][0] = t[12]
clf = clf.fit(d)
lab = clf.labels_
num = [0, 0, 0]
cen = clf.cluster_centers_
cen = [float(i) for i in cen]
t = cen.copy()
t.sort()
d = {}
for i, k1 in enumerate(t):
    for j, k2 in enumerate(cen):
        if k1 == k2:
        d[i] = j
break
for i, t in enumerate(compinf):
    compinf[i].append(int(d[lab[i]]))
num[d[lab[i]]] += 1
labels = ["A", "B", "C"]
plt.pie(x=num, labels=labels, autopct="%0.2f%%")
plt.title('第一问各企业分类饼状图')
plt.show()
# 确定三类企业的各自利率
cost = loadxls(3, 0, 2)
bestrate = [0, 0, 0]
bestpro = 0
for rate1 in range(0, 27):
    for rate2 in range(rate1 + 1, 28):
        for rate3 in range(rate2 + 1, 29):
        rate = [rate1, rate2, rate3]
sumpro = 0
for i in compinf:
    if round(i[2]) == 1:
        continue
sumpro += i[15] * cost[rate[i[16]]][0] * (1 - cost[rate[i[16]]][3 - round(3 * i[2])])
if sumpro > bestpro:
    bestpro = sumpro
rate = [cost[i][0] for i in rate]
bestrate = rate
print('最佳利率为:', bestrate)
print('最多获利为:', bestpro)
d = []
for i in compinf:
    if i[2] == 1:
        e = 0
l = 0
else:
e = i[15]
l = bestrate[2 - i[16]]
d.append([i[0], i[1], e, l])
savexls(compinf, '附件1的离散化信息',
        ['公司编号', '公司名称', '信用等级', '是否违约', '经营类型(税额)', '净收益', '总支出', '总收入',
         '交易取消率',
         '上游客户集中度', '下游客户集中度', '企业发展趋势', '信贷风险', '归一化风险', '最大额度',
         '实际额度', '风险聚类'])
savexls(d, '问题一贷款策略', ['公司编号', '公司名称', '实际额度', '贷款利率'])


# 问题二： 决策树代码
def ques2_1():
    comp_inf1 = loadxls(7, 0, 2, 2)


comp_inf2 = loadxls(8, 0, 2)
grades = 'D', 'C', 'B', 'A'
all_f = ['broken', 'business type', 'profit', 'outcome', 'income', 'transaction cancellation
         rate',
             'upstream influence', 'downstream influence', 'develop tendency']
data = np.array(comp_inf1)
for i, t, in enumerate(data):
    data[i][0] = np.round(t[0] * 3)


def gettree(cho, maxstep, minsplit, flag):
    d = data[:, cho]


f = []
for i in cho:
    f.append(all_f[i - 1])
# 生成训练集与测试集
# fr, to = 92, 106
# c = list(range(0, fr + 1))
# c.extend(list(range(to - 1, 123)))
# xtes = d[c, :]
# ytes = data[c, 0]
# xtra, ytra = d, data[:, 0]
xtra, xtes, ytra, ytes = train_test_split(d, data[:, 0], test_size=0.2)
clf = tree.DecisionTreeClassifier(criterion='entropy'
                                  , max_depth=maxstep
                                  , min_samples_split=minsplit
                                  )
clf.fit(xtra, ytra)
# 决策树示意图
file = tree.export_graphviz(clf, out_file=None, feature_names=f, class_names=grades,
                            filled=True, rounded=True, special_characters=True)
f = pydotplus.graph_from_dot_data(file)
f.write_pdf('tree' + str(flag) + '.pdf')
print('各个特征的影响力:', clf.feature_importances_)
# 训练集验证
answer = clf.predict(xtra)
ytra = ytra.reshape(-1)
num = 0
corr = 0
realcorr = 0
dincorr = 0
dnum = 0
for i, t in enumerate(answer):
    num += 1
if t == 3:
    dnum += 1
if t != ytra[i]:
    dincorr += 1
if abs(t - ytra[i]) <= 1:
    corr += 1
if abs(t - ytra[i]) == 0:
    realcorr += 1
print('训练集准确度:', corr / num)
if dnum != 0:
    t = dincorr / dnum
else:
    t = 0
print('训练集D等信誉错误率:', t)
print('训练集平均偏差:', np.mean(np.abs(answer - ytra)) / 3)
# 测试集验证
answer = clf.predict(xtes)
ytes = ytes.reshape(-1)
num = 0
corr = 0
dincorr = 0
dnum = 0
for i, t in enumerate(answer):
    if t == 3:
        dnum += 1
if t != ytes[i]:
    dincorr += 1
num += 1
if abs(t - ytes[i]) <= 1:
    corr += 1
print('测试集准确度:', corr / num)
if dnum != 0:
    t = dincorr / dnum
else:
    t = 0
print('测试集D等信誉错误率:', t)
print('测试集平均偏差:', np.mean(np.abs(answer - ytes)) / 3)
return clf
# 初步实现决策树
cho1 = [2, 3, 4, 5, 6, 7, 8, 9]
clf = gettree(cho1, maxstep=6, minsplit=10, flag=1)
# 降维简化数据
rer = []
rep = []
for col in range(2, 10):
    r, p = stats.pearsonr(data[:, 0], data[:, col])
rer.append(r)
rep.append(p)
print('各列相关度为:', end='')
for i in rer:
    print('%.3f' % (i), end='\t')
print(' ')
print('各列显著性为:', end='')
for i in rep:
    print('%.3f' % (i), end='\t')
print(' ')
# 剔除
cho2 = [4, 5, 6, 7, 8, 9]
gettree(cho2, maxstep=4, minsplit=15, flag=2)
# 预测
d = []
for i in comp_inf2:
    t = []
for j in i[2:]:
    if j == '':
        t.append(0)
else:
    t.append(float(j))
d.append(t)
data = np.array(d)
d = data[:, cho1]
ans = clf.predict(d)
for i, t in enumerate(ans):
    comp_inf2[i][2] = t / 3
comp_inf2[i][3] = 0
savexls(comp_inf2, '附件2的离散化信息',
        ['公司编号', '公司名称', '信用等级', '是否违约', '经营类型(税额)', '净收益', '总支出', '总收入',
         '交易取消率',
         '上游客户集中度', '下游客户集中度', '企业发展趋势', '信贷风险', '归一化风险', '最大额度',
         '实际额度', '风险聚类'])


def ques2_2():
    comp_inf = loadxls(8, 0, 2)


prof = loadxls(6, 0, 2)
# 计算风险与最大额度
wei = [0.29, 0, 0.06, -0.12, -0.12, -0.12, 0, 0.145, 0.145, -0.12]
mi = 1
ma = 0
for i, t1 in enumerate(comp_inf):
    sco = 0
for k, t2 in enumerate(t1[2:12]):
    sco += wei[k] * t2
if sco > ma:
    ma = sco
if sco < mi:
    mi = sco
comp_inf[i][12] = sco
for i, t1 in enumerate(comp_inf):
    comp_inf[i][13] = (1 - (comp_inf[i][12] - mi) / (ma - mi))
comp_inf[i][14] = float(prof[i + 123][-1])
comp_inf[i][15] = comp_inf[i][14] * float(comp_inf[i][13])
if comp_inf[i][15] > 1000000:
    comp_inf[i][15] = 1000000
elif comp_inf[i][15] < 100000:
    comp_inf[i][15] = 100000
# 风险聚类
clf = KMeans(n_clusters=3)
d = np.zeros((len(comp_inf), 1))
for i, t in enumerate(comp_inf):
    d[i][0] = t[12]
clf = clf.fit(d)
lab = clf.labels_
num = [0, 0, 0]
cen = clf.cluster_centers_
cen = [float(i) for i in cen]
t = cen.copy()
t.sort()
d = {}
for i, k1 in enumerate(t):
    for j, k2 in enumerate(cen):
        if k1 == k2:
        d[i] = j
break
for i, t in enumerate(comp_inf):
    comp_inf[i][16] = (int(d[lab[i]]))
num[d[lab[i]]] += 1
labels = ["A", "B", "C"]
plt.pie(x=num, labels=labels, autopct="%0.2f%%")
plt.title('第二问各企业分类饼状图')
plt.show()
# 确定三类企业的各自利率
cost = loadxls(3, 0, 2)
bestrate = [0, 0, 0]
bestpro = 0
bestmon = 0
for rate1 in range(0, 27):
    for rate2 in range(rate1 + 1, 28):
        for rate3 in range(rate2 + 1, 29):
        rate = [rate1, rate2, rate3]
sumpro = 0
summon = 0
for i in comp_inf:
    if round(i[2]) == 1:
        continue
sumpro += i[15] * cost[rate[i[16]]][0] * (1 - cost[rate[i[16]]][3 - round(3 * i[2])])
summon += i[15] * (1 - cost[rate[i[16]]][3 - round(3 * i[2])])
if summon <= 1e8 and sumpro > bestpro:
    bestpro = sumpro
rate = [cost[i][0] for i in rate]
bestrate = rate
bestmon = summon
print('最佳利率为:', bestrate)
print('最多获利为:', bestpro)
print('交易贷款总额为:', bestmon)
d = []
for i in comp_inf:
    if i[2] == 1:
        e = 0
l = 0
else:
e = i[15]
l = bestrate[2 - i[16]]
d.append([i[0], i[1], e, l])
savexls(d, '问题二贷款策略', ['公司编号', '公司名称', '实际额度', '贷款利率'])
savexls(comp_inf, '附件2的离散化信息',
        ['公司编号', '公司名称', '信用等级', '是否违约', '经营类型(税额)', '净收益', '总支出', '总收入',
         '交易取消率',
         '上游客户集中度', '下游客户集中度', '企业发展趋势', '信贷风险', '归一化风险', '最大额度',
         '实际额度', '风险聚类'])


# 计算考虑疫情影响和政府扶持下的各企业分数
def ques3_1(b):
    comp_inf = loadxls(8, 0, 2)


n = len(comp_inf[0])
clus = [['-0.5', '劳务', '个体经营', '影', '五金', '设计服务', '策划', '设计', '传播', '广告',
         '印务', '餐饮', '服务', '图书', '经营',
         '童装', '纺织品', '生活用品', '鞋', '服饰', '体育', '家居', '食品', '工贸', '文化', '贸易'],
        ['-0.25', '物资', '运', '物流', '汽贸', '快递', '商贸'],
        ['0', '建筑', '管理', '维修', '教育', '质量', '律师', '事务', '实业', '环保', '地质', '灾害',
         '土地', '生态', '猕猴桃', '园艺',
         '园林', '石化', '天然气', '装饰', '农业', '调味品', '蔬菜', '建设', '房地产', '建材', '包装',
         '材', '木业', '纸业', '塑胶',
         '卫浴', '门窗', '设备', '空调', '家电', '器材', '电器', '机电', '花', '轮胎', '化工', '机械',
         '车', '工程', '电气', '合金',
         '塑料'],
        ['0.25', '投资', '代理'],
        ['0.5', '发展', '科学', '电子', '技术', '网络', '医疗', '药', '科技']
        ]
for i, t in enumerate(comp_inf):
    for x in range(22 - len(comp_inf[i])):
        comp_inf[i].append(0)
flag = False
for k, c in enumerate(clus):
    if flag:
        break
for j in c:
    if j in t[1]:
        comp_inf[i][17] = float(clus[k][0])
comp_inf[i][19] = (b * comp_inf[i][17] - comp_inf[i][13] + 0.5) / (b + 1) + 0.5
comp_inf[i][18] = (b * comp_inf[i][17] + comp_inf[i][13] - 0.5) / (b + 1) + 0.5
comp_inf[i][20] = comp_inf[i][14] * float(comp_inf[i][19])
if comp_inf[i][20] > 1000000:
    comp_inf[i][20] = 1000000
elif comp_inf[i][20] < 100000:
    comp_inf[i][20] = 100000
flag = True
break
# 风险聚类
clf = KMeans(n_clusters=5)
d = np.zeros((len(comp_inf), 1))
for i, t in enumerate(comp_inf):
    d[i][0] = t[18]
clf = clf.fit(d)
lab = clf.labels_
num = [0, 0, 0, 0, 0]
cen = clf.cluster_centers_
cen = [float(i) for i in cen]
t = cen.copy()
t.sort()
d = {}
for i, k1 in enumerate(t):
    for j, k2 in enumerate(cen):
        if k1 == k2:
        d[i] = j
break
for i, t in enumerate(comp_inf):
    comp_inf[i][21] = (int(d[lab[i]]))
num[d[lab[i]]] += 1
labels = ["A", "B", "C", "D", "E"]
plt.pie(x=num, labels=labels, autopct="%0.2f%%")
plt.title('第三问各企业分类饼状图')
plt.show()
# 确定三类企业的各自利率
cost = loadxls(3, 0, 2)
bestrate = [0, 0, 0, 0, 0]
bestpro = 0
bestmon = 0
for rate1 in range(0, 25):
    for rate2 in range(rate1 + 1, 26):
        for rate3 in range(rate2 + 1, 27):
        for rate4 in range(rate3 + 1, 28):
        for rate5 in range(rate4 + 1, 29):
        rate = [rate1, rate2, rate3, rate4, rate5]
sumpro = 0
summon = 0
for i in comp_inf:
    if round(i[2]) == 1:
        continue
sumpro += i[20] * cost[rate[i[21]]][0] * (1 - cost[rate[i[21]]][3 - round(3 * i[2])])
summon += i[20] * (1 - cost[rate[i[21]]][3 - round(3 * i[2])])
if summon <= 1e8 and sumpro > bestpro:
    bestpro = sumpro
rate = [cost[i][0] for i in rate]
bestrate = rate
bestmon = summon
print('最佳利率为:', bestrate)
print('最多获利为:', bestpro)
print('交易贷款总额为:', bestmon)
d = []
for i in comp_inf:
    if i[2] == 1:
        e = 0
l = 0
else:
e = i[20]
l = bestrate[4 - i[21]]
d.append([i[0], i[1], e, l])
if b == 2:
    savexls(d, '问题三贷款策略', ['公司编号', '公司名称', '实际额度', '贷款利率'])
savexls(comp_inf, '附件2的离散化信息',
        ['公司编号', '公司名称', '信用等级', '是否违约', '经营类型(税额)', '净收益', '总支出', '总收入',
         '交易取消率', '上游客户集中度',
         '下游客户集中度', '企业发展趋势', '信贷风险', '归一化风险', '最大额度', '实际额度', '风险聚类',
         '影响值', '分数1', '分数2',
         '考虑疫情的额度', '分数聚类'])
return bestpro
ques1_1()
ques1_2()
ques1_3()
ques1_4()
ques2_1()
ques2_2()
ques3_1(2)
# 灵敏度分析
i = 1
lmd = []
x = []
while (i <= 3):
    print(i)
x.append(i)
lmd.append(ques3_1(i))
i += 0.125
plt.plot(x, lmd)
plt.xlabel('疫情影响权重比')
plt.ylabel('银行最大获利')
plt.title('疫情影响权重比与银行最大获利的灵敏度分析曲线图')
plt.show()
