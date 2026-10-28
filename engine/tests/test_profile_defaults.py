#!/usr/bin/env python
"""学校画像 + 按画像兜底默认值（v0.15 新增）。跑 python test_profile_defaults.py。"""

import json
import os
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
ENGINE = os.path.dirname(HERE)
sys.path.insert(0, ENGINE)

WORK = tempfile.mkdtemp(prefix='tk-profile-')
os.environ['TEACHER_KIT_PROFILE'] = os.path.join(WORK, 'profile.json')

from teacher_kit import defaults, profile as store  # noqa: E402
from teacher_kit.commands import TABLE  # noqa: E402

FAILS = []


def check(condition, label, detail=''):
    if condition:
        print('  ✓ ' + label)
    else:
        FAILS.append(label)
        print('  ✗ ' + label + (' — ' + str(detail) if detail else ''))


FULL = {
    'school': {'name': '临河区第三中学', 'stage': '初中', 'grade': '七年级', 'subject': '英语',
               'teacher': '张老师', 'classNames': '七（1）班', 'schoolYear': '2026 春'},
    'paper': {'size': '8K', 'orientation': 'landscape', 'columns': 2,
              'footer': '使用班级：{{className}}　出卷老师：{{teacherName}}', 'note': '听力先放'},
    'plan': {'layout': '表格式', 'note': '一课时一页'},
    'deck': {'palette': '学校蓝', 'fontSize': 28, 'slideSize': '16:9'},
    'wechat': {'tone': '正式', 'signature': '临河区第三中学德育处'},
    'docflow': {'unit': '临河区第三中学', 'dept': '德育处', 'prefix': '临三中'},
    'tasks': {'minutes': 90, 'answerCard': True, 'columns': 2},
}


def main():
    print('— 洗数据：脏东西不许进画像')
    clean, ignored = store.sanitize({
        'school': {'name': ' 临河区第三中学 ', 'bad-key!': 'x', 'v': 3},
        'nonsense': {'a': 1},
        'paper': {'footer': 'x' * 500},
        'plan': 'not-a-dict',
    })
    check(ignored == ['nonsense'], '不认识的节被忽略（不报错）', ignored)
    check(clean['school']['name'] == '临河区第三中学', '首尾空白被剪掉', clean['school'])
    check('bad-key!' not in clean['school'], '非法键名丢掉', clean['school'])
    check(clean['school'].get('v') == 3, '整数照样能存', clean['school'])
    check(len(clean['paper']['footer']) == 200, '超长值截到 200 字',
          len(clean['paper']['footer']))
    check('plan' not in clean, '该是对象的位置给了字符串 → 整节丢掉', clean)

    print('— 存 / 读 / 记学到什么')
    saved = store.update(FULL, note='学校信息和卷面规格记下来了', task='profile', learn=True)
    check(saved['saved'] == ['deck', 'docflow', 'paper', 'plan', 'school', 'tasks', 'wechat'],
          '七节都存进去了', saved['saved'])
    check(os.path.isfile(saved['path']), '画像落在指定路径', saved['path'])
    check(not os.path.exists(saved['path'] + '.tmp'), '临时文件不留在盘上')
    loaded, path, exists = store.load()
    check(exists is True, '再读 exists 为真')
    check(loaded['school']['name'] == '临河区第三中学', '校名读回来了')
    check(len(loaded['history']) == 1, '记了一条 history', len(loaded['history']))
    check(loaded['history'][0]['note'] == '学校信息和卷面规格记下来了',
          'history 带那句给人看的备注', loaded['history'][0])
    check(loaded['history'][0]['task'] == 'profile', 'history 带 task')
    check(len(loaded['learned']) == 1, 'learn=True 时记一条 learned')
    check(len(loaded['history'][0].get('keys') or []) > 5, 'history 记了哪些键变了')

    store.update({'plan': {'note': '两栏灰线'}}, note='教案页边距改了一下')
    again, _path, _exists = store.load()
    check(len(again['learned']) == 1, 'learn=False 不往 learned 里加')
    check(len(again['history']) == 2, 'history 继续累加', len(again['history']))
    check(again['school']['name'] == '临河区第三中学', '后一次更新不清掉之前记的')

    print('— 纸型两写法都要认')
    check(store.paper_key_of({'size': '8K', 'orientation': 'landscape', 'columns': 2}) == '8K-2',
          '老师写法（8K 横向两栏）认成 8K-2')
    check(store.paper_key_of({'size': '8开', 'columns': 2}) == '8K-2', '中文「8开」也认')
    check(store.paper_key_of({'key': 'A4-1'}) == 'A4-1', '内部名直接命中')
    check(store.paper_key_of({'size': 'A3', 'orientation': 'portrait'}) == 'A3-1',
          '竖排 → 一栏')
    check(store.paper_page({'size': '不存在'}) is None, '认不出纸型 → paper_page 给 None（不瞎猜）')
    check(store.paper_key_of({'size': '不存在'}) not in store.PAPER_LABELS, '认不出就不折成内部名')
    page = store.paper_page({'size': '8K', 'orientation': 'landscape', 'columns': 2})
    check(page == {'size': '8K', 'orientation': 'landscape', 'columns': 2},
          'paper_page 给 paper.resolve 认的形状', page)
    check(store.paper_page({'key': 'A4-1'})['orientation'] == 'portrait', 'A4 是竖向')

    print('— 说成人话')
    text = store.describe(loaded)
    check('临河区第三中学' in text and '初中七年级英语' in text, '头一行是学校与学科', text.splitlines()[0] if text else '')
    check('8 开横向两栏' in text, '纸型说成老师的话（不是 8K-2）', text)
    check('表格式' in text, '教案版式说成中文')
    check('公众号语气：正式' in text, '公众号语气进描述')
    check('落款 临河区第三中学德育处' in text, '公众号落款进描述')
    check('本单位 临河区第三中学' in text, '公文单位进描述')
    check('常用考试时长：90' in text, '常用时长进描述')
    check(text.startswith('· '), '每行以 · 开头（好往引导语里贴）')

    empty_text = store.describe(store.empty())
    check(empty_text == '', '空画像不产出半截描述（免得模型以为学校是空的）', repr(empty_text))

    print('— 给宿主/面板的概要')
    brief = store.summary()
    check(brief['exists'] is True, 'exists 为真')
    check(brief['text'] == store.describe(store.load()[0]), '概要里的 text 就是那段人话')
    check(isinstance(brief['learnedCount'], int) and brief['learnedCount'] == 1, 'learnedCount 是数字')
    check('history' not in brief['profile'] and 'learned' not in brief['profile'],
          '概要里不带 history（省得刷屏）', sorted(brief['profile']))
    check(brief['filled'] and 'school' in brief['filled'], 'filled 列出记住的节', brief['filled'])
    check(brief['path'] == os.environ['TEACHER_KIT_PROFILE'], 'path 指着隔离的画像文件')

    print('— 画像坏了不许拖垮生成')
    with open(os.environ['TEACHER_KIT_PROFILE'], 'w', encoding='utf-8') as handle:
        handle.write('{ 这不是 JSON')
    broken, _path, exists = store.load()
    check(exists is False, '读坏了就当没有（exists False）')
    check(store.describe(broken) == '', '读坏了 describe 也只是空字符串')
    recovered = store.update({'school': {'name': '临河区第三中学'}}, note='重来一次')
    check(recovered['profile']['school']['name'] == '临河区第三中学', '坏文件能被覆盖修好')
    json.load(open(os.environ['TEACHER_KIT_PROFILE'], encoding='utf-8'))

    print('— 按画像兜底：只补没说到的')
    store.update(FULL, note='恢复完整画像')
    profile = store.load()[0]

    exam = defaults.apply({'text': '一、听力'}, 'exam:build', profile)
    check(exam.get('page') == {'size': '8K', 'orientation': 'landscape', 'columns': 2},
          '没给纸型 → 用学校的 8 开横向两栏', exam.get('page'))
    check(exam['values']['teacherName'] == '张老师', '出卷老师用画像里的')
    check(exam['values']['className'] == '七（1）班', '班级用画像里的')
    check(exam['options']['minutes'] == 90, '考试时长用画像里的', exam['options'])
    check(exam['options']['answerCard'] == 'card', '画像要答题卡 → 默认带卡')
    check('page=8Kx2' in exam['_profileFilled'], '记下补了哪些项', exam.get('_profileFilled'))

    keep = defaults.apply({'page': {'size': 'A4', 'orientation': 'portrait', 'columns': 1},
                           'values': {'teacherName': '李老师'},
                           'options': {'minutes': 45}}, 'exam:build', profile)
    check(keep['page']['size'] == 'A4', '老师这次指定了纸型 → 画像不许覆盖', keep['page'])
    check(keep['values']['teacherName'] == '李老师', '指定了老师姓名 → 不覆盖')
    check(keep['options']['minutes'] == 45, '指定了时长 → 不覆盖')

    tpl = defaults.apply({'text': '一、听力', 'baseTemplate': __file__}, 'exam:build', profile)
    check('page' not in tpl, '套模板时不硬塞 page（模板比画像权威）', tpl.get('page'))
    check(tpl.get('_profilePage', {}).get('size') == '8K', '但留一份画像纸型当兜底', tpl.get('_profilePage'))

    plan = defaults.apply({'text': '一、教学内容'}, 'plan:build', profile)
    check(plan['layout'] == 'form', '画像里的「表格式」→ 引擎认的 form', plan.get('layout'))
    check(plan['teacher'] == '张老师', '教案署名用画像里的')

    deck = defaults.apply({'text': '一、导入'}, 'lesson:deck', profile)
    check(deck['options']['bodySizePt'] == 28, '课件正文字号用画像里的', deck['options'])
    check(deck['options']['slideSize'] == '16:9', '画布比例用画像里的')
    check(not deck['options'].get('titleSizePt'), '画像没记标题字号就不瞎填')

    wechat = defaults.apply({'blocks': ['一段话']}, 'wechat:render', profile)
    check(wechat['signature'] == '临河区第三中学德育处', '公众号落款用画像里的')

    docflow = defaults.apply({'title': '关于 X 的通知'}, 'docflow:build', profile)
    check(docflow['unit'] == '临河区第三中学', '公文本单位用画像里的')
    check(docflow['dept'] == '德育处', '公文承办科室用画像里的')

    other = defaults.apply({'text': 'x'}, 'docx:clean', profile)
    check('_profileFilled' not in other and 'page' not in other,
          '跟画像无关的命令一点不动', other)

    nobody = defaults.apply({'text': '一、听力'}, 'exam:build', store.empty())
    check('page' not in nobody, '画像空就不补（不许凭空造 8 开）', nobody)

    print('— 清空')
    cleared = store.reset()
    check(cleared['profile']['school'] == {}, 'reset 清掉学校')
    check(store.load()[2] is True, 'reset 之后文件还在（只是空的）')
    check(store.describe(store.load()[0]) == '', 'reset 之后描述为空')

    print('— 命令表接线')
    check(callable(TABLE.get('profile:get')), 'profile:get 在命令表里')
    check(callable(TABLE.get('profile:set')), 'profile:set 在命令表里')
    check(callable(TABLE.get('profile:reset')), 'profile:reset 在命令表里')

    print()
    if FAILS:
        print('失败 ' + str(len(FAILS)) + ' 项：')
        for label in FAILS:
            print(' - ' + label)
        sys.exit(1)
    print('全部通过 ✅')


if __name__ == '__main__':
    main()
