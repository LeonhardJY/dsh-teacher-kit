
import { readFileSync } from 'node:fs'
import vm from 'node:vm'
import { join } from 'node:path'
const root = "D:\\eepseek Harness Workspace\\teacher-kit"
const bundle = readFileSync(join(root, 'client', 'client.js'), 'utf8')
let definition
const calls = []
const sandbox = {
  window: { __ModuleLoader__: { load: (def) => { definition = def } } },
  console,
  fetch: (url, init) => { calls.push(url); return Promise.resolve({ ok: true, json: () => Promise.resolve({ ok: true }) }) }
}
vm.runInNewContext(bundle, sandbox, { filename: 'client.js' })

function makeReact() {
  const state = { slots: [], deps: {}, cursor: 0 }
  let pending = null
  let queue = []
  let dirty = false
  function schedule(kind, fn, deps) {
    const i = state.cursor++
    const key = kind + i
    const prev = state.deps[key]
    const changed = !prev || !deps || deps.length !== prev.length || deps.some((d, k) => !Object.is(d, prev[k]))
    if (!changed) return
    state.deps[key] = deps ? deps.slice() : null
    queue.push(fn)
  }
  const api = {
    createElement(type, props, ...children) {
      const kids = children.length === 0 ? undefined : children.length === 1 ? children[0] : children
      return { type, props: Object.assign({}, props || {}, { children: kids }) }
    },
    useState(initial) {
      const i = state.cursor++
      if (!(i in state.slots)) state.slots[i] = initial
      const set = (value) => {
        const next = typeof value === 'function' ? value(state.slots[i]) : value
        if (Object.is(next, state.slots[i])) return
        state.slots[i] = next
        dirty = true
      }
      return [state.slots[i], set]
    },
    useMemo(fn) { return fn() },
    useCallback(fn) { return fn },
    useLayoutEffect(fn, deps) { schedule('l', fn, deps) },
    useEffect(fn, deps) { schedule('e', fn, deps) }
  }
  api.render = (component, props) => {
    const draw = () => {
      state.cursor = 0
      queue = []
      const tree = component(props)
      for (const fn of queue.splice(0)) fn()
      if (dirty) { dirty = false; return draw() }
      return tree
    }
    pending = () => { draw() }
    return draw()
  }
  api.slots = () => state.slots.slice()
  return api
}
const React = makeReact()
const mod = definition.factory((id) => { if (id === 'react') return React; throw new Error(id) })
let captured
let injected
const ctx = {
  slots: {
    inject: (name, factory) => { injected = factory },
    register: (entry, component) => { captured = { entry, component } }
  },
  logger: { warn: () => {} }
}
mod.apply(ctx)
injected()
const C = captured.component
const r1 = React.render(C, { onActiveChange: () => {}, sessionId: 'sA' })
console.log('r1', r1.type, 'slots', React.slots())
r1.props.onClick && r1.props.onClick()
console.log('after click slots', React.slots(), 'calls', calls)
const r2 = React.render(C, { onActiveChange: () => {}, sessionId: 'sA' })
console.log('r2', r2.type)
const r3 = React.render(C, { onActiveChange: () => {}, sessionId: 'sB' })
console.log('r3', r3.type)
