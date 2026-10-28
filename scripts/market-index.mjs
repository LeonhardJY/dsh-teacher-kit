
import { execSync } from 'node:child_process';
import { readFileSync, writeFileSync } from 'node:fs';

const commit = execSync('git rev-parse HEAD', { cwd: 'D:/eepseek Harness Workspace/teacher-kit' }).toString().trim();
const icon = readFileSync('market/icon-dataurl.txt', 'utf8').trim();
const pkg = JSON.parse(readFileSync('package.json', 'utf8'));
const now = Math.floor(Date.now() / 1000);

const index = {
  format: 1,
  title: { zh: '教师助手', en: 'Teacher Kit' },
  description: {
    zh: '面向中国大陆中学教师的 DSH 插件索引：出试卷、组卷、听写、教案、课件、整套、文件加工与教务辅助。',
    en: 'DSH plugin index for middle-school teachers in China: exams, papers, dictation, lesson plans, slide decks, lesson packs, docx work and school-office helpers.'
  },
  metadata: {
    categories: {
      education: { title: { zh: '教育', en: 'Education' } }
    },
    labels: {
      docx: { title: { zh: 'Word 文档', en: 'Word documents' } },
      pptx: { title: { zh: 'PPT 课件', en: 'Slide decks' } },
      'local-first': { title: { zh: '全本地', en: 'Local-first' } }
    },
    sortKeys: {
      order: { title: { zh: '推荐位', en: 'Featured order' }, order: 'asc' }
    }
  },
  plugins: [
    {
      name: pkg.name,
      title: { zh: '教师助手', en: 'Teacher Kit' },
      description: {
        zh: '给中国大陆中学老师的办公助手：出试卷、按题库组卷、听写单、教案、课件、一节课整套，以及 docx/pptx 批量加工。排版走你自己学校的空白模板，第一次用问几个问题，之后记住你学校的习惯。全程本机运行，不联网、不收集数据。',
        en: "An office assistant for middle-school teachers in China: exams, question-bank papers, dictation sheets, lesson plans, slide decks and full lesson packs, plus docx/pptx batch jobs. Lays everything out on your own school template and remembers your school's habits after the first run. Runs entirely on your machine — no network, no data collection."
      },
      icon,
      author: 'GoldenUniverse',
      homepage: 'https://github.com/LeonhardJY/dsh-teacher-kit',
      repository: { url: 'https://github.com/LeonhardJY/dsh-teacher-kit.git' },
      metadata: {
        category: 'education',
        keywords: ['教师', '试卷', '教案', '课件', 'docx', 'pptx', 'education', 'teacher', 'exam'],
        labels: ['docx', 'pptx', 'local-first'],
        sortKeys: { order: 1 }
      },
      versions: [
        {
          version: pkg.version,
          publishTimestamp: now,
          source: {
            type: 'git',
            url: 'https://github.com/LeonhardJY/dsh-teacher-kit.git',
            commit
          },
          engines: { dsh: pkg.engines.dsh }
        }
      ]
    }
  ]
};

writeFileSync('market/index.json', JSON.stringify(index, null, 2) + '\n');
console.log('commit', commit, 'version', pkg.version, 'bytes', JSON.stringify(index).length);
