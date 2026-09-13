import json, re
with open(r'D:\Desktop\外国新闻事业史1\news-history-app\js\data.js', 'r', encoding='utf-8') as f:
    content = f.read()
match = re.search(r'const APP_DATA = (\{.*?\});', content, re.DOTALL)
if match:
    data = json.loads(match.group(1))
    print('数据加载成功！')
    print('  背诵主题:', len(data['materials']), '个')
    print('  背诵知识点:', data['stats']['totalMaterials'], '条')
    print('  题库题目:', data['stats']['totalQuestions'], '题')
    print('    单选:', len(data['questions']['single']))
    print('    多选:', len(data['questions']['multiple']))
    print('    名词解释:', len(data['questions']['term']))
    print('    简答:', len(data['questions']['short']))
    print('    论述:', len(data['questions']['essay']))
    print('  复习计划:', len(data['plan']), '天')
    print('  主题列表:', list(data['materials'].keys()))
else:
    print('数据解析失败！')
