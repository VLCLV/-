import os

app_dir = r'D:\Desktop\外国新闻事业史1\news-history-app'
pages = ['index.html','today.html','calendar.html','materials.html',
         'keypoints.html','quiz.html','mistakes.html','exam.html']

# 移动端优化meta标签
mobile_metas = '''  <meta name="apple-mobile-web-app-capable" content="yes">
  <meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
  <meta name="apple-mobile-web-app-title" content="新闻史复习">
  <meta name="mobile-web-app-capable" content="yes">
  <meta name="theme-color" content="#1F3864">
  <meta name="format-detection" content="telephone=no">'''

for page in pages:
    path = os.path.join(app_dir, page)
    with open(path, 'r', encoding='utf-8') as f:
        content = f.read()

    # 检查是否已经添加过
    if 'apple-mobile-web-app-capable' in content:
        print(f'{page}: 已优化，跳过')
        continue

    # 在viewport标签后插入移动端meta
    viewport_tag = '<meta name="viewport" content="width=device-width, initial-scale=1.0">'
    if viewport_tag in content:
        content = content.replace(viewport_tag, viewport_tag + '\n' + mobile_metas)
        with open(path, 'w', encoding='utf-8') as f:
            f.write(content)
        print(f'{page}: 已添加移动端优化')
    else:
        print(f'{page}: 未找到viewport标签，跳过')

print('\n全部完成！')
