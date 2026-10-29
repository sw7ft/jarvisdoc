/* The guide is the one part of JarvisDoc that argues for a change of habit, which is
   exactly the kind of thing software has historically got wrong. The rules it plays by:
   it never speaks while you are typing, at most twice in a sitting, never repeats a
   lesson, and stops volunteering altogether once three lessons go unread. The avatar
   stays put either way, so the whole argument is still there when you go looking for it. */

const STORE = 'jd.guide'

const QUIET_AFTER_TYPING = 12000
const SETTLE_AFTER_LAUNCH = 90000
const SETTLE_AFTER_OPEN = 25000
const BETWEEN_LESSONS = 8 * 60 * 1000
const PER_SESSION = 2
const GIVE_UP_AFTER = 3
const READ_TIME = 6000
const LINGER = 26000

function kb(bytes) {
  if (!bytes) return '0 KB'
  if (bytes < 1024) return `${bytes} bytes`
  return `${Math.round(bytes / 102.4) / 10} KB`
}

function utf8Bytes(text) {
  return new TextEncoder().encode(text || '').length
}

/* Every lesson makes one argument and stops. `when` keeps a lesson back until the open
   document actually demonstrates it, so the guide is never caught claiming something the
   evidence on screen contradicts. */
const LESSONS = [
  {
    id: 'permanence',
    title: 'Your words outlive the program',
    body:
      'A <code>.docx</code> is a zip file full of XML. Reading one means owning software that still understands that zip, that XML, and that particular version of it. ' +
      'A <code>.md</code> file is the letters you typed, in the order you typed them. Every text editor written in the last fifty years can open it, and the ones written in the next fifty will too.'
  },
  {
    id: 'plaintext',
    title: 'You can read it without us',
    body:
      'Open this document in TextEdit or Notepad right now and the first thing you will see is <code>PK</code> followed by a wall of binary. ' +
      'Open the Markdown version and you will see your own sentences, with a <code>#</code> in front of the headings. ' +
      'If JarvisDoc disappeared tomorrow, one of those two files still works.',
    action: 'show',
    actionLabel: 'Show me'
  },
  {
    id: 'size',
    title: 'Most of a .docx is not your document',
    body:
      'The scaffolding inside a Word file — style tables, font lists, relationship maps, a content-type manifest — travels with every copy you send, forever. ' +
      'Markdown carries the words and nothing else.',
    when: (ctx) => ctx.format === 'docx' && ctx.fileBytes > 0 && ctx.markdownBytes > 0 && ctx.fileBytes > ctx.markdownBytes * 2,
    evidence: (ctx) =>
      `This file is <strong>${kb(ctx.fileBytes)}</strong> on disk. Everything you actually wrote is <strong>${kb(ctx.markdownBytes)}</strong> of Markdown. ` +
      `The remaining ${kb(ctx.fileBytes - ctx.markdownBytes)} is scaffolding.`
  },
  {
    id: 'diff',
    title: 'You can see exactly what changed',
    body:
      'Version history can show you what a Word document said, but it cannot point at the three words someone altered in paragraph nine. ' +
      'Markdown is lines, so any tool can line two versions up and mark the difference. ' +
      'That is the reason every engineering team on earth keeps its documents this way.'
  },
  {
    id: 'structure',
    title: 'It will not let you fake a heading',
    body:
      'In Word a heading is often just bigger, bolder text. It looks like a heading and behaves like nothing. ' +
      'Markdown has no way to pretend: <code>##</code> is a heading or it is not one. ' +
      'That single constraint is what makes tables of contents, outlines, and screen readers work.'
  },
  {
    id: 'search',
    title: 'Everything can search inside it',
    body:
      'Spotlight, <code>grep</code>, your backup system and your editor can all look inside a Markdown file without unpacking anything first. ' +
      'A <code>.docx</code> has to be decompressed and parsed before any of them can find a phone number in it.'
  },
  {
    id: 'lockin',
    title: 'Nobody owns Markdown',
    body:
      'The <code>.docx</code> format is a specification controlled by one company, and it has changed before. ' +
      'Markdown is a convention. There is no licence to renew, no subscription that can put your own sentences behind a paywall, and no release that quietly stops opening last decade’s files.'
  },
  {
    id: 'everywhere',
    title: 'It pastes into everything',
    body:
      'GitHub, Slack, Reddit, Discord, Notion, Obsidian, every static site generator and most wikis already read Markdown. ' +
      'The document stops being a file you attach and becomes text you can put anywhere, still formatted.'
  },
  {
    id: 'machines',
    title: 'Machines read it without a translator',
    body:
      'Markdown is what language models read and write natively. ' +
      'A plain text document can be summarised, searched, rewritten or published as a web page with nothing in between — no export step, no conversion, no fidelity argument.'
  },
  {
    id: 'merge',
    title: 'It survives sync, backup and two authors',
    body:
      'Two people editing the same <code>.docx</code> produces two files and a conversation about which one is current. ' +
      'Two people editing the same Markdown file produces a change that can actually be merged. ' +
      'And a decade of versions costs less disk space than one copy of the Word file.'
  }
]

const RECAP = {
  id: 'recap',
  title: 'That is the whole argument',
  body:
    'Permanence, plain text, real diffs, honest structure, no owner. ' +
    'Nothing here asks you to abandon Word today — JarvisDoc keeps writing <code>.docx</code> for as long as you need it. ' +
    'The offer is only ever made when a document would survive the trip intact.'
}

function readState() {
  try {
    const saved = JSON.parse(localStorage.getItem(STORE) || '{}')
    return {
      learned: saved.learned && typeof saved.learned === 'object' ? saved.learned : {},
      ignored: Number(saved.ignored) || 0,
      stopped: Boolean(saved.stopped)
    }
  } catch {
    return { learned: {}, ignored: 0, stopped: false }
  }
}

const guide = {
  state: readState(),
  enabled: true,
  deps: {},
  el: {},
  ctx: { format: 'docx', fileBytes: 0, markdownBytes: 0, words: 0, ready: false },
  current: null,
  shownAt: 0,
  engaged: false,
  spokeAt: 0,
  spokenThisSession: 0,
  lastTypedAt: 0,
  startedAt: Date.now(),
  openedAt: 0,
  pending: false,
  timer: 0,
  idleTimer: 0,

  save() {
    try {
      localStorage.setItem(STORE, JSON.stringify(this.state))
    } catch {
      /* a full disk should never take the editor down with it */
    }
  },

  mount(deps) {
    this.deps = deps || {}
    const id = (name) => document.getElementById(name)
    this.el = {
      root: id('guide'),
      bubble: id('guide-bubble'),
      title: id('guide-title'),
      body: id('guide-body'),
      proof: id('guide-proof'),
      progress: id('guide-progress'),
      avatar: id('guide-avatar'),
      badge: id('guide-badge'),
      act: id('guide-do'),
      next: id('guide-next'),
      stop: id('guide-stop'),
      close: id('guide-x')
    }
    if (!this.el.root) return

    this.el.avatar.addEventListener('click', () => {
      if (this.isOpen()) this.close()
      else this.summon()
    })
    this.el.close.addEventListener('click', () => this.close())
    this.el.next.addEventListener('click', () => {
      this.engaged = true
      this.summon()
    })
    this.el.act.addEventListener('click', () => {
      this.engaged = true
      const action = this.current?.action
      this.close()
      if (action === 'save') this.deps.saveAsMarkdown?.()
      else this.deps.showMarkdown?.()
    })
    this.el.stop.addEventListener('click', () => {
      this.state.stopped = true
      this.save()
      this.close()
      this.deps.toast?.('The guide will stay quiet — click it any time')
    })

    this.paintBadge()
  },

  setEnabled(on) {
    this.enabled = Boolean(on)
    if (!this.enabled) this.close()
    this.el.root?.toggleAttribute('hidden', !this.enabled)
  },

  isOpen() {
    return Boolean(this.el.bubble) && !this.el.bubble.hidden
  },

  remaining() {
    return LESSONS.filter((lesson) => !this.state.learned[lesson.id])
  },

  paintBadge() {
    if (!this.el.badge) return
    const show = !this.isOpen() && (this.pending || Object.keys(this.state.learned).length === 0)
    this.el.badge.hidden = !show
    this.el.avatar?.classList.toggle('nudging', this.pending && !this.isOpen())
  },

  /* Context arrives from whatever the editor happens to know; the guide only ever reads it. */
  note(event, data) {
    if (event === 'typing') {
      this.lastTypedAt = Date.now()
      return
    }
    if (event === 'report' && data) {
      this.ctx.format = data.format || this.ctx.format
      this.ctx.ready = Boolean(data.ready)
      this.ctx.fileBytes = Number(data.fileBytes) || 0
      this.ctx.markdownBytes = utf8Bytes(data.markdown)
    }
    if (event === 'opened' && data) {
      this.ctx.format = data.format || 'docx'
      this.openedAt = Date.now()
      this.close()
    }
    if (event === 'words') this.ctx.words = Number(data) || 0
    this.consider()
  },

  // Nothing is said until the document is worth talking about and the room is quiet.
  consider() {
    if (!this.enabled || this.isOpen()) return
    if (this.ctx.format !== 'docx' || this.ctx.words < 60) return
    if (!this.pick()) return
    this.pending = true
    this.paintBadge()
    if (this.state.stopped) return
    if (this.spokenThisSession >= PER_SESSION) return
    if (Date.now() - this.startedAt < SETTLE_AFTER_LAUNCH) return
    if (this.openedAt && Date.now() - this.openedAt < SETTLE_AFTER_OPEN) return
    if (this.spokeAt && Date.now() - this.spokeAt < BETWEEN_LESSONS) return
    this.waitForPause()
  },

  waitForPause() {
    window.clearTimeout(this.idleTimer)
    const quietFor = Date.now() - this.lastTypedAt
    if (quietFor < QUIET_AFTER_TYPING) {
      this.idleTimer = window.setTimeout(() => this.waitForPause(), QUIET_AFTER_TYPING - quietFor + 250)
      return
    }
    if (!this.enabled || this.isOpen() || this.state.stopped) return
    if (this.deps.busy?.()) return
    const lesson = this.pick()
    if (!lesson) return
    this.spokenThisSession += 1
    this.spokeAt = Date.now()
    this.show(lesson, { auto: true })
  },

  pick() {
    return this.remaining().find((lesson) => !lesson.when || lesson.when(this.ctx)) || null
  },

  summon() {
    this.engaged = true
    this.show(this.pick() || RECAP, { auto: false })
  },

  show(lesson, { auto }) {
    if (!this.el.bubble) return
    this.settleCurrent()
    this.current = lesson
    this.shownAt = Date.now()
    this.engaged = !auto
    this.pending = false

    this.el.title.textContent = lesson.title
    this.el.body.innerHTML = lesson.body

    const proof = lesson.evidence?.(this.ctx)
    this.el.proof.hidden = !proof
    if (proof) this.el.proof.innerHTML = proof

    // The migration offer only appears where it is honestly true.
    const canMove = this.ctx.ready && this.ctx.format === 'docx'
    const action = canMove ? 'save' : lesson.action
    this.el.act.hidden = !action
    if (action) this.el.act.textContent = canMove ? 'Save as Markdown' : lesson.actionLabel || 'Show me'
    if (action) this.current = { ...lesson, action }

    const place = Object.keys(this.state.learned).length + 1
    this.el.progress.textContent = lesson.id === 'recap' ? 'all read' : `${Math.min(place, LESSONS.length)} of ${LESSONS.length}`
    this.el.next.hidden = lesson.id === 'recap'

    this.el.bubble.hidden = false

    if (lesson.id !== 'recap') {
      this.state.learned[lesson.id] = Date.now()
      this.save()
    }
    this.paintBadge()

    window.clearTimeout(this.timer)
    if (auto) this.timer = window.setTimeout(() => this.close(), LINGER)
  },

  // A lesson that was shut instantly did not land. Three of those and the guide gives up.
  settleCurrent() {
    if (!this.current || !this.shownAt) return
    const read = this.engaged || Date.now() - this.shownAt > READ_TIME
    this.state.ignored = read ? 0 : this.state.ignored + 1
    if (this.state.ignored >= GIVE_UP_AFTER) this.state.stopped = true
    this.save()
    this.shownAt = 0
  },

  close() {
    if (!this.el.bubble) return
    this.settleCurrent()
    this.el.bubble.hidden = true
    this.current = null
    window.clearTimeout(this.timer)
    this.paintBadge()
  }
}

window.JarvisGuide = guide
