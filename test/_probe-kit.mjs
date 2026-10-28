import { pathToFileURL } from 'node:url'
const target = 'D:/Deepseek Harnees/resources/app.asar.unpacked/dsh/node_modules/@deepseek-ai/libreoffice-kit/lib/index.js'
try {
  const mod = await import(pathToFileURL(target).href)
  console.log('keys', Object.keys(mod).join(','), typeof mod.createConverter)
} catch (e) { console.log('FAIL', e.message.slice(0, 300)) }
console.log('VIPSHOME', process.env.VIPSHOME)
