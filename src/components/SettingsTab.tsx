import { useEffect, useMemo, useRef, useState } from 'react'
import { createPortal } from 'react-dom'
import { api } from '../ipc'
import { type Office, officeScopedKey, TEMPLATE_KEYS } from '../office'
import { corporationByName } from '../zoneCircleDirectory'
import {
  WORK_ORDER_TEMPLATE_VARIANTS,
  DEFAULT_WORK_ORDER_TEMPLATE_VARIANT,
  AGREEMENT_TEMPLATE_VARIANTS,
  DEFAULT_AGREEMENT_TEMPLATE_VARIANT,
  INTIMATION_TEMPLATE_VARIANTS,
  DEFAULT_INTIMATION_TEMPLATE_VARIANT,
  FILE_BACKER_TEMPLATE_VARIANTS,
  DEFAULT_FILE_BACKER_TEMPLATE_VARIANT,
  CIVIL_TENDER_TEMPLATE_VARIANTS,
  DEFAULT_CIVIL_TENDER_TEMPLATE_VARIANT,
  type TemplateVariantOption
} from '@core/workOrderTemplateVariants'
import {
  workOrderPlaceholders,
  agreementPlaceholders,
  fileBackerPlaceholders,
  civilTenderPlaceholders,
  type WorkOrderAgreementFields
} from '@core/workOrderAgreement'
import { intimationPlaceholders } from '@core/intimationFill'
import { MAX_CONCURRENT_SESSIONS } from '@core/sessionSlots'
import { THEME_OPTIONS, type ThemeId } from '../theme'
import type { ActiveSessionInfo } from '../../electron/ipc-contract'
import { base64ToUint8, renderDocPreview } from './docPage'
import { IconSettings, IconCheck, IconWarn, IconOpen, IconUser, IconLogout, IconChevronRight } from './Icons'

interface Props {
  office: Office
  theme: ThemeId
  onThemeChange: (theme: ThemeId) => void
}

// Deliberately long/realistic values (not blank, not short) — a preview built
// from these surfaces the same class of alignment/overflow problem real data
// would, the way it did while building the Kompally variant (a too-wide
// amounts block only showed up once filled with real-length figures). Circle/
// CNO/Zone/Corporation come from the office actually selected (falling back to
// a placeholder office when none is picked yet) — showing a tile filled with
// another office's circle name (e.g. Kompally) while a different office is
// selected reads as if the template itself belongs to that other circle.
function sampleFields(office: Office): WorkOrderAgreementFields {
  const circle = office.circle || 'Sample Circle'
  const cno = office.circleNumber || '00'
  const zone = office.zone || 'Sample'
  const corporation = office.corporation || 'CMC'
  const corporationFullName = corporationByName(corporation)?.fullName || 'Cyberabad Municipal Corporation'
  return {
    circle,
    cno,
    zone,
    nameOfWork: `Laying of CC Road Block No 80 to 82 Near Rajiv Gruha Kalpa in Ward No-292 Sai Baba Nagar in ${circle} Circle-${cno}, ${zone} Zone, ${corporation}`,
    itemNo: '1',
    agencyName: 'Sri Venkateswara Infrastructure and Constructions Private Limited',
    address: 'H.No.5-24-347, Srinivas Nagar, Gajularamaram, Quthbullapur, Hyderabad, Telangana -500055',
    phone: '9876543210',
    wincode: 'WC-2026-00456',
    financialYear: '2026-27',
    estimateLakhs: '11.90',
    ecvRupees: '943549',
    tenderPercent: '21.60',
    contractRupees: '739742.42',
    workOrderDate: '',
    agreementDate: '',
    adminSanctionDate: '09.02.2026',
    corporation,
    corporationFullName,
    tsNoDate: '',
    ceLetterNoDate: '',
    completionMonths: '03',
    reservation: '',
    emdDetails: '',
    noticeNo: `01/EE-${cno}/QBZ/${corporation}/2025-2026`,
    tenderId: '679920',
    noticeDate: '09-02-2026',
    intimationDate: '17.04.2026'
  }
}

// Settings' preview tile just needs some representative placeholder values —
// no real L-1 upload or Agreement/Schedule-A extras to read them from — so
// the bid-date/contractor fields resolve to the sample office's own values
// and the two hand-entered extras (pages of Agreement, Schedule 'A' items)
// stay blank the same way an unfilled document would show them.
function civilTenderSettingsPlaceholders(f: WorkOrderAgreementFields): Record<string, string> {
  return civilTenderPlaceholders(f, {}, {})
}

interface TemplateSectionProps {
  title: string
  subtitle: string
  storageKey: string
  variants: TemplateVariantOption[]
  defaultVariant: string
  fields: WorkOrderAgreementFields
  fetchTemplate: (variantId: string) => Promise<string>
  placeholders: (f: WorkOrderAgreementFields) => Record<string, string>
}

/**
 * One document type's tile row — the same "render a real thumbnail, click to
 * act" gallery already used for the Work Order/Agreement documents themselves
 * (WorkOrderAgreementTab's .wo-tile), filled with realistic sample data so
 * each tile shows exactly how that variant actually looks, not just its
 * name. Clicking a tile selects it; the selected one is marked. The small eye
 * button opens the same LibreOffice-accurate render at full size (same
 * "wo-modal" preview WorkOrderAgreementTab's own document tiles use), since a
 * distorted thumbnail alone can't convince anyone it matches the real
 * document. Self-contained (its own selection/preview state) so a second
 * document type is just another instance of this, per office.
 */
function TemplateSection({ title, subtitle, storageKey, variants, defaultVariant, fields, fetchTemplate, placeholders }: TemplateSectionProps) {
  const [selected, setSelected] = useState(() => localStorage.getItem(storageKey) || defaultVariant)
  const [failed, setFailed] = useState<Set<string>>(new Set())
  const [previewId, setPreviewId] = useState<string | null>(null)
  const [open, setOpen] = useState(false)
  const tileRefs = useRef(new Map<string, HTMLDivElement>())
  const modalRef = useRef<HTMLDivElement | null>(null)
  const selectedVariant = variants.find((v) => v.id === selected)

  function choose(id: string) {
    setSelected(id)
    localStorage.setItem(storageKey, id)
  }

  async function fillVariant(id: string): Promise<string> {
    const b64 = await fetchTemplate(id)
    const labels = await api.findPlaceholdersInDocument(b64)
    const values = placeholders(fields)
    const resolved = labels.map((label) => ({ label, column: label, score: 1 }))
    return api.fillPlaceholdersInDocument(b64, resolved, values)
  }

  // Render every variant's live thumbnail whenever the office (and so the
  // sample data) changes — filled with the LibreOffice-accurate render, not
  // the fast approximate one. Only a couple of tiles render per section (not
  // the 11+ a document catalog can have), so the conversion cost that rules
  // out the accurate renderer elsewhere doesn't apply. One variant's render
  // failing (a missing/corrupt bundled file) only greys out that one tile,
  // never blocks the others.
  useEffect(() => {
    let cancelled = false
    for (const v of variants) {
      void (async () => {
        try {
          const filled = await fillVariant(v.id)
          if (cancelled) return
          const container = tileRefs.current.get(v.id)
          if (!container) return
          await renderDocPreview(base64ToUint8(filled), container)
        } catch {
          if (!cancelled) setFailed((prev) => new Set(prev).add(v.id))
        }
      })()
    }
    return () => {
      cancelled = true
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [fields])

  // The full-size, same accurate render — re-rendered into the modal at full
  // size whenever a tile's preview is opened, so the entire document (every
  // page) can actually be read, not just the cropped tile thumbnail.
  useEffect(() => {
    if (!previewId) return
    const container = modalRef.current
    if (!container) return
    let cancelled = false
    container.innerHTML = ''
    void fillVariant(previewId)
      .then((filled) => {
        if (cancelled) return
        return renderDocPreview(base64ToUint8(filled), container)
      })
      .catch(() => {
        if (!cancelled) setFailed((prev) => new Set(prev).add(previewId))
      })
    return () => {
      cancelled = true
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [previewId])

  const previewVariant = variants.find((v) => v.id === previewId)

  return (
    <div className="settings-template-section">
      <div className="settings-section-head">
        <button
          type="button"
          className="settings-section-toggle"
          onClick={() => setOpen((o) => !o)}
          aria-expanded={open}
        >
          <IconChevronRight className={`settings-section-chevron ${open ? 'open' : ''}`} />
          <span className="settings-section-toggle-titles">
            <h3 className="settings-template-section-title">{title}</h3>
            <p className="sub">{subtitle}</p>
          </span>
        </button>
        <div className="settings-section-head-actions">
          {selectedVariant && <span className="settings-section-current">{selectedVariant.label}</span>}
          <button
            type="button"
            className="settings-section-preview-btn"
            title={`Preview ${selectedVariant?.label ?? title}`}
            onClick={() => setPreviewId(selected)}
          >
            <IconOpen />
          </button>
        </div>
      </div>

      <div className={`wo-tiles settings-template-tiles ${open ? '' : 'settings-section-collapsed'}`}>
        {variants.map((v) => (
          <div
            key={v.id}
            role="button"
            tabIndex={0}
            className={`wo-tile ${selected === v.id ? 'on' : ''}`}
            onClick={() => choose(v.id)}
            onKeyDown={(e) => {
              if (e.key === 'Enter' || e.key === ' ') {
                e.preventDefault()
                choose(v.id)
              }
            }}
            title={`Use "${v.label}" as this office's ${title}`}
          >
            <div className="wo-tile-preview">
              {failed.has(v.id) ? (
                <div className="settings-tile-error">
                  <IconWarn /> Couldn't load
                </div>
              ) : (
                <div ref={(el) => void (el ? tileRefs.current.set(v.id, el) : tileRefs.current.delete(v.id))} className="wo-tile-doc" />
              )}
              {selected === v.id && (
                <span className="settings-tile-selected">
                  <IconCheck /> In use
                </span>
              )}
              <button
                type="button"
                className="wo-tile-open"
                title={`See the full ${v.label} document`}
                onClick={(e) => {
                  e.stopPropagation()
                  setPreviewId(v.id)
                }}
              >
                Click to preview
              </button>
            </div>
            <div className="wo-tile-foot">{v.label}</div>
          </div>
        ))}
      </div>

      {previewId &&
        createPortal(
          <div className="wo-modal-overlay" onClick={() => setPreviewId(null)}>
            <div className="wo-modal" onClick={(e) => e.stopPropagation()}>
              <div className="wo-modal-head">
                <span className="wo-modal-title">{previewVariant?.label}</span>
                <button className="wo-modal-close" onClick={() => setPreviewId(null)} title="Close" aria-label="Close">
                  ×
                </button>
              </div>
              <div className="wo-modal-body">
                {failed.has(previewId) ? (
                  <div className="settings-tile-error">
                    <IconWarn /> Couldn't load this document.
                  </div>
                ) : (
                  <div ref={modalRef} className="intimation-docx-preview" />
                )}
              </div>
            </div>
          </div>,
          document.body
        )}
    </div>
  )
}

/** "just now" / "12 min ago" / "3 hr ago" / a plain date once it's more than a day old. */
function relativeTime(ts: number): string {
  const mins = Math.round((Date.now() - ts) / 60000)
  if (mins < 1) return 'just now'
  if (mins < 60) return `${mins} min ago`
  const hours = Math.round(mins / 60)
  if (hours < 24) return `${hours} hr ago`
  return new Date(ts).toLocaleDateString()
}

/**
 * Settings — "Active Devices": every device currently signed in as this
 * login ID (up to MAX_CONCURRENT_SESSIONS — see core/sessionSlots.ts), with a
 * "Log out" button for every device except this one, so someone locked out
 * of a stuck old session (or who just wants to free a slot) doesn't have to
 * physically walk to the other device. Ending THIS device's own session
 * deliberately isn't offered here — that's the existing profile-menu Logout,
 * which also tears down this device's local listener; this list only ever
 * releases an OTHER device's slot (electron/main.ts's logoutOtherSession
 * handler refuses to do otherwise).
 */
function ActiveDevicesCard() {
  const [sessions, setSessions] = useState<ActiveSessionInfo[] | null>(null)
  const [loadError, setLoadError] = useState(false)
  const [loggingOut, setLoggingOut] = useState<string | null>(null)

  async function refresh() {
    try {
      const list = await api.listActiveSessions()
      setSessions(list)
      setLoadError(false)
    } catch {
      setLoadError(true)
    }
  }

  useEffect(() => {
    void refresh()
  }, [])

  async function logoutDevice(session: ActiveSessionInfo) {
    const label = session.deviceLabel || 'this device'
    if (!window.confirm(`Sign "${label}" out? It will need to log in again to continue working.`)) return
    setLoggingOut(session.sessionId)
    try {
      await api.logoutOtherSession(session.sessionId)
      await refresh()
    } finally {
      setLoggingOut(null)
    }
  }

  return (
    <div className="card">
      <div className="card-head">
        <div className="head-ic">
          <IconUser />
        </div>
        <div className="titles">
          <h2>Active Devices</h2>
          <p className="sub">
            Every device currently signed in with this login (up to {MAX_CONCURRENT_SESSIONS} at a time). Log out a
            device you don't recognize or no longer use to free up a slot.
          </p>
        </div>
      </div>

      {sessions === null && !loadError && <p className="hint">Loading…</p>}
      {loadError && (
        <div className="notice error">
          <IconWarn /> Could not load active devices — check your connection and try again.
        </div>
      )}
      {sessions && sessions.length === 0 && <p className="hint">No active devices found.</p>}

      {sessions && sessions.length > 0 && (
        <ul className="active-devices-list">
          {sessions.map((s) => (
            <li key={s.sessionId} className="active-device-row">
              <div className="active-device-info">
                <span className="active-device-label">
                  {s.deviceLabel || 'Unknown device'}
                  {s.isThisDevice && <span className="active-device-tag">This device</span>}
                </span>
                <span className="active-device-meta">
                  Signed in {relativeTime(s.loginAt)} · last active {relativeTime(s.lastSeenAt)}
                </span>
              </div>
              {!s.isThisDevice && (
                <button
                  type="button"
                  className="active-device-logout"
                  disabled={loggingOut === s.sessionId}
                  onClick={() => void logoutDevice(s)}
                >
                  <IconLogout /> {loggingOut === s.sessionId ? 'Logging out…' : 'Log out'}
                </button>
              )}
            </li>
          ))}
        </ul>
      )}
    </div>
  )
}

/**
 * Settings — "Themes": lets a user try the flat/solid-color Issue Documents
 * tile style ("Test theme 1", built as the .doc-tile-flat CSS variant) side
 * by side with the original thumbnail tiles ("Default"), and switch back at
 * any time. A static mock tile (not a real DocThumbnail render) — this card
 * doesn't have a real .docx on hand to render, and the point here is
 * comparing the two tile *styles*, not previewing actual document content.
 * This is a display preference for this machine, not office data, so it's
 * stored in localStorage only (see theme.ts) rather than office-scoped.
 */
function ThemeSection({ theme, onThemeChange }: { theme: ThemeId; onThemeChange: (theme: ThemeId) => void }) {
  return (
    <div className="settings-template-section">
      <div className="wo-tiles settings-theme-tiles">
        {THEME_OPTIONS.map((opt) => (
          <div
            key={opt.id}
            role="button"
            tabIndex={0}
            className={`wo-tile ${theme === opt.id ? 'on' : ''}`}
            onClick={() => onThemeChange(opt.id)}
            onKeyDown={(e) => {
              if (e.key === 'Enter' || e.key === ' ') {
                e.preventDefault()
                onThemeChange(opt.id)
              }
            }}
            title={`Use "${opt.label}" for Issue Documents tiles`}
          >
            <div className="wo-tile-preview settings-theme-preview">
              <div
                className={`doc-tile-card tone-sky settings-theme-mock ${
                  opt.id === 'flat1'
                    ? 'doc-tile-flat'
                    : opt.id === 'windows'
                      ? 'settings-theme-mock-windows'
                      : opt.id === 'dark'
                        ? 'settings-theme-mock-dark'
                        : ''
                }`}
              >
                {opt.id === 'flat1' ? (
                  <span className="doc-tile-flat-thumb">
                    <span className="settings-theme-mock-thumb" />
                  </span>
                ) : (
                  <span className="settings-theme-mock-thumb" />
                )}
                <span className="doc-tile-card-name">Sample Document</span>
                <span className="doc-tile-card-meta">Added 01.01.2026</span>
              </div>
              {theme === opt.id && (
                <span className="settings-tile-selected">
                  <IconCheck /> In use
                </span>
              )}
              <span className="wo-tile-open">Click to use this style</span>
            </div>
            <div className="wo-tile-foot">{opt.label}</div>
            <p className="settings-theme-desc">{opt.description}</p>
          </div>
        ))}
      </div>
    </div>
  )
}

/**
 * Settings — "Document Templates": some circles word and lay out their own
 * documents differently (see core/workOrderTemplateVariants.ts for why — a
 * real Kompally Circle-56 Work Order cites a Ref block the app's original
 * template has no room for at all, and its own Agreement Bond is a full
 * "AGREEMENT - DEED" with contractor legal boilerplate the app's default
 * cover page doesn't carry). Every variant is a real circle's own template,
 * gathered and bundled with the app — never a runtime upload — so picking one
 * here just tells the app which bundled file this office's document should
 * use, stored per office the same way as every other office-scoped setting
 * (see CONTACT_KEYS). One TemplateSection per document type — a future
 * document type is just another entry in the list below.
 */
export default function SettingsTab({ office, theme, onThemeChange }: Props) {
  const fields = useMemo(
    () => sampleFields(office),
    [office.circle, office.circleNumber, office.zone, office.corporation]
  )
  const [themesOpen, setThemesOpen] = useState(false)
  const currentTheme = THEME_OPTIONS.find((o) => o.id === theme)

  return (
    <>
      <ActiveDevicesCard />

      <div className="card">
        <button
          type="button"
          className="card-head settings-section-toggle settings-card-head-toggle"
          onClick={() => setThemesOpen((o) => !o)}
          aria-expanded={themesOpen}
        >
          <div className="head-ic">
            <IconSettings />
          </div>
          <span className="settings-section-toggle-titles">
            <h2>Themes</h2>
            <p className="sub">Try a different look for this app on this device.</p>
          </span>
          <div className="settings-section-head-actions">
            {currentTheme && <span className="settings-section-current">{currentTheme.label}</span>}
            <IconChevronRight className={`settings-section-chevron ${themesOpen ? 'open' : ''}`} />
          </div>
        </button>

        {themesOpen && <ThemeSection theme={theme} onThemeChange={onThemeChange} />}
      </div>

      <div className="card">
        <div className="card-head">
          <div className="head-ic">
            <IconSettings />
          </div>
          <div className="titles">
            <h2>Document Templates</h2>
            <p className="sub">Click a style below to make it this office's default. Each tile is a live preview filled with sample data, not just a name.</p>
          </div>
        </div>

        <TemplateSection
          title="Work Order"
          subtitle="Some circles word and lay out their own Work Order differently."
          storageKey={officeScopedKey(TEMPLATE_KEYS.workOrder, office)}
          variants={WORK_ORDER_TEMPLATE_VARIANTS}
          defaultVariant={DEFAULT_WORK_ORDER_TEMPLATE_VARIANT}
          fields={fields}
          fetchTemplate={api.workOrderTemplate}
          placeholders={workOrderPlaceholders}
        />

        <TemplateSection
          title="Agreement Bond"
          subtitle="Some circles word and lay out their own Agreement Bond differently."
          storageKey={officeScopedKey(TEMPLATE_KEYS.agreement, office)}
          variants={AGREEMENT_TEMPLATE_VARIANTS}
          defaultVariant={DEFAULT_AGREEMENT_TEMPLATE_VARIANT}
          fields={fields}
          fetchTemplate={api.agreementTemplate}
          placeholders={agreementPlaceholders}
        />

        <TemplateSection
          title="Intimation"
          subtitle="Some circles word and lay out their own Intimation letter differently."
          storageKey={officeScopedKey(TEMPLATE_KEYS.intimation, office)}
          variants={INTIMATION_TEMPLATE_VARIANTS}
          defaultVariant={DEFAULT_INTIMATION_TEMPLATE_VARIANT}
          fields={fields}
          fetchTemplate={api.intimationTemplate}
          placeholders={intimationPlaceholders}
        />

        <TemplateSection
          title="File Backer"
          subtitle="Some circles word and lay out their own File Backer cover page differently."
          storageKey={officeScopedKey(TEMPLATE_KEYS.fileBacker, office)}
          variants={FILE_BACKER_TEMPLATE_VARIANTS}
          defaultVariant={DEFAULT_FILE_BACKER_TEMPLATE_VARIANT}
          fields={fields}
          fetchTemplate={api.fileBackerTemplate}
          placeholders={fileBackerPlaceholders}
        />

        <TemplateSection
          title="Tender Document"
          subtitle="Some circles word and lay out their own Tender Document header differently."
          storageKey={officeScopedKey(TEMPLATE_KEYS.civilTender, office)}
          variants={CIVIL_TENDER_TEMPLATE_VARIANTS}
          defaultVariant={DEFAULT_CIVIL_TENDER_TEMPLATE_VARIANT}
          fields={fields}
          fetchTemplate={api.civilTenderTemplate}
          placeholders={civilTenderSettingsPlaceholders}
        />
      </div>
    </>
  )
}
