import { useEffect, useMemo, useRef, useState } from 'react'
import {
  Check,
  Download,
  FileSpreadsheet,
  KeyRound,
  LogOut,
  Plus,
  Save,
  Search,
  Trash2,
  Upload,
  UserRound,
} from 'lucide-react'
import JSZip from 'jszip'
import { saveAs } from 'file-saver'
import * as XLSX from 'xlsx'
import './App.css'

const DEMO_LOGIN = 'admin'
const DEMO_PASSWORD = '12345'
const APP_VERSION = 'restore-excel-after-refresh-2026-05-01-0026'
const CLOUD_SAVE_URL = import.meta.env.VITE_CLOUD_SAVE_URL
const WORK_DRAFT_KEY = 'police-helper-work-draft'
const DOCUMENT_EDITS_DRAFT_KEY = 'police-helper-document-edits-draft'
const RECENT_VIEWER_FILES_KEY = 'police-helper-recent-viewer-files'
const RECENT_HELPER_FILES_KEY = 'police-helper-recent-helper-files'

const TARGET_TYPES = {
  actNumber: {
    label: 'Номер акта',
    marker: 'Номер акта',
  },
  erdrExtract: {
    label: 'Витяг з ЄРДР/ЄО',
    marker: 'Витяг з ЄРДР/ЄО',
  },
}

const createId = () => {
  if (globalThis.crypto?.randomUUID) {
    return globalThis.crypto.randomUUID()
  }

  return `id-${Date.now()}-${Math.random().toString(16).slice(2)}`
}

const createRule = (targetType = 'actNumber') => ({
  id: createId(),
  targetType,
  personKeys: [],
  day: '',
  value: '',
})

const normalize = (value) =>
  String(value ?? '')
    .trim()
    .toLowerCase()

const isLikelyName = (value) => {
  const text = String(value ?? '').trim()
  if (text.length < 2 || text.length > 80) return false
  if (/[0-9=]/.test(text)) return false
  return /^[A-Za-zА-Яа-яЁёІіЇїЄєҐґ'\-\s.]+$/.test(text)
}

const getColumnLetter = (index) => XLSX.utils.encode_col(index)

const getColumnByDay = (day) => {
  const dayNumber = Number.parseInt(day, 10)
  if (!Number.isInteger(dayNumber) || dayNumber < 1 || dayNumber > 30) return ''

  return getColumnLetter(XLSX.utils.decode_col('G') + dayNumber - 1)
}

const DOCUMENT_EDIT_FIELDS = [
  { key: 'number', label: 'Номер' },
  { key: 'date', label: 'Дата' },
  {
    key: 'agency',
    label: 'Найменування органу (підрозділу) поліції, який видав документ',
  },
  {
    key: 'signer',
    label: 'Прізвище, власне ім’я, по батькові (за наявності), посада особи, що підписала',
  },
]

const MAY_2026_DATES = Array.from({ length: 31 }, (_, index) => {
  const day = String(index + 1).padStart(2, '0')
  return `${day}.05.2026`
})

const EO_AGENCY_OPTIONS = [
  'ВП № 1 Краматорського РУП ГУНП в Донецькій області',
  'Краматорський РУП ГУНП в Донецькій області',
  'ВП № 3 Краматорського РУП ГУНП в Донецькій області',
]

const ACT_SIGNER_OPTIONS = [
  'Начальник УВТС ГУНП в Донецькій області, полковник поліції Сергій НЕСТЕРОВ',
  'Заступник начальника УВТС ГУНП в Донецькій області, підполковник поліції Сергій КОЖЕДУБ',
  'Старший інспектор ВЗВ № 1 УВТС ГУНП в Донецькій області, капітан поліції Максим БІЛЕНКО',
  'Старший інспектор ВЗВ № 1 УВТС ГУНП в Донецькій області, старший лейтенант поліції Богдан ПІТЕРІН',
  'Старший інспектор з ОД ВЗВ № 1 УВТС ГУНП в Донецькій області, капітан поліції Віталій КУЗНЄЦОВ',
]

const EO_SIGNER_OPTIONS = [
  'Начальник ВП № 1 Краматорського РУП ГУНП в Донецькій області, підполковник поліції Євгеній ТУГАЙ',
  'Т.в.о. начальника Краматорського РУП ГУНП в Донецькій області, підполковник поліції Артем КУЗНЄЦОВ',
  'Т.в.о. начальника ВП № 3 Краматорського РУП ГУНП в Донецькій області, полковник поліції Ігор УГНІВЕНКО',
]

const findHeaderColumn = (row, matcher) =>
  row.findIndex((cell) => matcher(normalize(cell)))

const findDocumentTable = (rows) => {
  const headerIndex = rows.findIndex((row) => {
    const rowText = row.map((cell) => normalize(cell)).join(' ')
    return (
      rowText.includes('вид документа') &&
      rowText.includes('номер') &&
      rowText.includes('дата')
    )
  })

  if (headerIndex === -1) {
    return { rows: [], columns: {} }
  }

  const headerRow = rows[headerIndex] ?? []
  const columns = {
    type: findHeaderColumn(headerRow, (value) => value.includes('вид документа')),
    number: findHeaderColumn(headerRow, (value) => value === 'номер'),
    date: findHeaderColumn(headerRow, (value) => value === 'дата'),
    agency: findHeaderColumn(
      headerRow,
      (value) => value.includes('найменування') && value.includes('орган'),
    ),
    signer: findHeaderColumn(
      headerRow,
      (value) => value.includes('прізвище') && value.includes('підписала'),
    ),
  }

  if (columns.type < 0) columns.type = 0
  if (columns.number < 0) columns.number = 4
  if (columns.date < 0) columns.date = 6
  if (columns.agency < 0) columns.agency = 8
  if (columns.signer < 0) columns.signer = 17

  const documentRows = []
  for (let index = headerIndex + 1; index < rows.length; index += 1) {
    const row = rows[index] ?? []
    const documentType = String(row[columns.type] ?? '').trim()
    const hasDocumentValues = DOCUMENT_EDIT_FIELDS.some(({ key }) =>
      String(row[columns[key]] ?? '').trim(),
    )

    if (!documentType && !hasDocumentValues) {
      if (documentRows.length) break
      continue
    }

    if (!documentType) continue

    documentRows.push({
      rowNumber: index + 1,
      type: documentType,
      number: String(row[columns.number] ?? ''),
      date: String(row[columns.date] ?? ''),
      agency: String(row[columns.agency] ?? ''),
      signer: String(row[columns.signer] ?? ''),
    })
  }

  return { rows: documentRows, columns }
}

const findTargetRowsForPerson = (rows, startIndex) => {
  const targetRows = {}
  const maxIndex = Math.min(rows.length, startIndex + 35)
  let isTargetSection = false

  for (let index = startIndex; index < maxIndex; index += 1) {
    const rowText = (rows[index] ?? [])
      .map((cell) => String(cell ?? '').trim())
      .join(' ')
    if (rowText.includes('Винагорода до 100000')) {
      isTargetSection = true
    }

    if (!isTargetSection) continue

    Object.entries(TARGET_TYPES).forEach(([key, target]) => {
      if (!targetRows[key] && rowText.includes(target.marker)) {
        targetRows[key] = index + 1
      }
    })
  }

  return targetRows
}

const escapeXml = (value) =>
  String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;')

const escapeRegExp = (value) => value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')

const getColumnIndex = (address) =>
  XLSX.utils.decode_col(address.match(/[A-Z]+/i)?.[0] ?? 'A')

const arrayBufferToDataUrl = (buffer) =>
  new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = () => resolve(reader.result)
    reader.onerror = reject
    reader.readAsDataURL(new Blob([buffer]))
  })

const dataUrlToArrayBuffer = async (dataUrl) => {
  const response = await fetch(dataUrl)
  return response.arrayBuffer()
}

const loadRecentViewerFiles = () => {
  try {
    return JSON.parse(localStorage.getItem(RECENT_VIEWER_FILES_KEY) ?? '[]')
  } catch {
    localStorage.removeItem(RECENT_VIEWER_FILES_KEY)
    return []
  }
}

const loadRecentHelperFiles = () => {
  try {
    return JSON.parse(localStorage.getItem(RECENT_HELPER_FILES_KEY) ?? '[]')
  } catch {
    localStorage.removeItem(RECENT_HELPER_FILES_KEY)
    return []
  }
}

const buildSheetView = (sheet) => {
  if (!sheet?.['!ref']) {
    return { cells: [], columnWidths: [] }
  }

  const range = XLSX.utils.decode_range(sheet['!ref'])
  const merges = sheet['!merges'] ?? []
  const covered = new Set()
  const mergeStarts = new Map()

  merges.forEach((merge) => {
    const startAddress = XLSX.utils.encode_cell(merge.s)
    mergeStarts.set(startAddress, {
      colSpan: merge.e.c - merge.s.c + 1,
      rowSpan: merge.e.r - merge.s.r + 1,
    })

    for (let row = merge.s.r; row <= merge.e.r; row += 1) {
      for (let col = merge.s.c; col <= merge.e.c; col += 1) {
        const address = XLSX.utils.encode_cell({ r: row, c: col })
        if (address !== startAddress) covered.add(address)
      }
    }
  })

  const rowLimit = Math.min(range.e.r, range.s.r + 199)
  const colLimit = Math.min(range.e.c, range.s.c + 79)
  const cells = []

  for (let row = range.s.r; row <= rowLimit; row += 1) {
    const rowCells = []
    for (let col = range.s.c; col <= colLimit; col += 1) {
      const address = XLSX.utils.encode_cell({ r: row, c: col })
      if (covered.has(address)) continue

      const cell = sheet[address]
      rowCells.push({
        address,
        value: cell?.w ?? cell?.v ?? '',
        colSpan: mergeStarts.get(address)?.colSpan ?? 1,
        rowSpan: mergeStarts.get(address)?.rowSpan ?? 1,
      })
    }
    cells.push({
      height: sheet['!rows']?.[row]?.hpx,
      cells: rowCells,
    })
  }

  const columnWidths = []
  for (let col = range.s.c; col <= colLimit; col += 1) {
    columnWidths.push(sheet['!cols']?.[col]?.wpx ?? 86)
  }

  return { cells, columnWidths }
}

const getCellXml = (value) => {
  const normalizedValue = String(value ?? '').trim()
  const numericValue = normalizedValue.replace(',', '.')

  if (normalizedValue && /^-?\d+(?:[,.]\d+)?$/.test(normalizedValue)) {
    return {
      typeAttribute: '',
      valueXml: `<v>${numericValue}</v>`,
      sharedStringIndex: null,
    }
  }

  return {
    typeAttribute: ' t="s"',
    valueXml: null,
    sharedStringIndex: null,
  }
}

const appendSharedString = (sharedStringsXml, value) => {
  const stringXml = `<si><t>${escapeXml(value)}</t></si>`
  const siCount = (sharedStringsXml.match(/<si>/g) ?? []).length
  let nextXml = sharedStringsXml.replace('</sst>', `${stringXml}</sst>`)

  if (/\bcount="\d+"/.test(nextXml)) {
    nextXml = nextXml.replace(/\bcount="\d+"/, (match) => {
      const current = Number.parseInt(match.match(/\d+/)?.[0] ?? '0', 10)
      return `count="${current + 1}"`
    })
  }

  if (/\buniqueCount="\d+"/.test(nextXml)) {
    nextXml = nextXml.replace(/\buniqueCount="\d+"/, (match) => {
      const current = Number.parseInt(match.match(/\d+/)?.[0] ?? '0', 10)
      return `uniqueCount="${current + 1}"`
    })
  }

  return { xml: nextXml, index: siCount }
}

const getSheetXmlPath = async (zip, targetSheetName) => {
  const workbookXml = await zip.file('xl/workbook.xml')?.async('string')
  const relsXml = await zip.file('xl/_rels/workbook.xml.rels')?.async('string')
  if (!workbookXml || !relsXml) return null

  const sheetRegex = /<sheet\b[^>]*>/g
  const sheets = workbookXml.match(sheetRegex) ?? []
  const sheetTag = sheets.find((tag) => {
    const name = tag.match(/\bname="([^"]*)"/)?.[1]
    return name === targetSheetName
  })
  const relationshipId = sheetTag?.match(/\br:id="([^"]*)"/)?.[1]
  if (!relationshipId) return null

  const relRegex = new RegExp(
    `<Relationship\\b[^>]*\\bId="${escapeRegExp(relationshipId)}"[^>]*/?>`,
  )
  const relTag = relsXml.match(relRegex)?.[0]
  const target = relTag?.match(/\bTarget="([^"]*)"/)?.[1]
  if (!target) return null

  if (target.startsWith('/')) return target.replace(/^\/+/, '')
  return `xl/${target.replace(/^\.?\//, '')}`
}

const replaceCellValue = (sheetXml, address, valuePayload) => {
  const rowNumber = XLSX.utils.decode_cell(address).r + 1
  const valueXml =
    valuePayload.valueXml ?? `<v>${valuePayload.sharedStringIndex}</v>`
  const cellXml = `<c r="${address}"${valuePayload.typeAttribute}>${valueXml}</c>`
  const cellRegex = new RegExp(
    `<c\\b(?=[^>]*\\br="${escapeRegExp(address)}")[^>]*/>|<c\\b(?=[^>]*\\br="${escapeRegExp(address)}")[^>]*>[\\s\\S]*?<\\/c>`,
  )

  if (cellRegex.test(sheetXml)) {
    return sheetXml.replace(cellRegex, (existingCell) => {
      const attrs = existingCell.match(/^<c\b([^>]*)/)?.[1] ?? ''
      const keptAttrs = attrs
        .replace(/\s+t="[^"]*"/g, '')
        .replace(/\s+r="[^"]*"/g, '')
        .replace(/\/\s*$/g, '')
      return `<c r="${address}"${keptAttrs}${valuePayload.typeAttribute}>${valueXml}</c>`
    })
  }

  const rowRegex = new RegExp(
    `<row\\b(?=[^>]*\\br="${rowNumber}")[^>]*>[\\s\\S]*?<\\/row>`,
  )
  if (rowRegex.test(sheetXml)) {
    return sheetXml.replace(rowRegex, (rowXml) => {
      const cells =
        rowXml.match(/<c\b[^>]*\br="[^"]+"[^>]*\/>|<c\b[^>]*\br="[^"]+"[^>]*>[\s\S]*?<\/c>/g) ??
        []
      const nextCell = cells.find((cell) => {
        const cellAddress = cell.match(/\br="([^"]*)"/)?.[1]
        return cellAddress && getColumnIndex(cellAddress) > getColumnIndex(address)
      })

      if (nextCell) return rowXml.replace(nextCell, `${cellXml}${nextCell}`)
      return rowXml.replace('</row>', `${cellXml}</row>`)
    })
  }

  const newRowXml = `<row r="${rowNumber}">${cellXml}</row>`
  return sheetXml.replace('</sheetData>', `${newRowXml}</sheetData>`)
}

function App() {
  const [isAuthorized, setIsAuthorized] = useState(
    localStorage.getItem('excel-tool-auth') === 'yes',
  )
  const [login, setLogin] = useState('')
  const [password, setPassword] = useState('')
  const [authError, setAuthError] = useState('')
  const [workbook, setWorkbook] = useState(null)
  const [sourceBuffer, setSourceBuffer] = useState(null)
  const [fileName, setFileName] = useState('')
  const [sheetName, setSheetName] = useState('')
  const [nameColumn, setNameColumn] = useState('')
  const [searchTerm, setSearchTerm] = useState('')
  const [manualPeople, setManualPeople] = useState([])
  const [manualName, setManualName] = useState('')
  const [manualRow, setManualRow] = useState('')
  const [documentEdits, setDocumentEdits] = useState({})
  const [activeTargetType, setActiveTargetType] = useState('actNumber')
  const [rulesByTarget, setRulesByTarget] = useState({
    actNumber: [createRule('actNumber')],
    erdrExtract: [createRule('erdrExtract')],
  })
  const [activeRuleId, setActiveRuleId] = useState('')
  const [status, setStatus] = useState('')
  const [currentPage, setCurrentPage] = useState('helper')
  const [viewerWorkbook, setViewerWorkbook] = useState(null)
  const [viewerFileName, setViewerFileName] = useState('')
  const [viewerSheetName, setViewerSheetName] = useState('')
  const [viewerZoom, setViewerZoom] = useState(1)
  const [recentViewerFiles, setRecentViewerFiles] = useState(loadRecentViewerFiles)
  const [recentHelperFiles, setRecentHelperFiles] = useState(loadRecentHelperFiles)
  const [activeHelperRecentId, setActiveHelperRecentId] = useState('')
  const [activeViewerRecentId, setActiveViewerRecentId] = useState('')
  const fileInputRef = useRef(null)
  const viewerFileInputRef = useRef(null)

  const activeSheet = workbook && sheetName ? workbook.Sheets[sheetName] : null
  const rows = useMemo(() => {
    if (!activeSheet) return []
    return XLSX.utils.sheet_to_json(activeSheet, { header: 1, defval: '' })
  }, [activeSheet])

  const documentTable = useMemo(() => findDocumentTable(rows), [rows])
  const documentRows = useMemo(
    () =>
      documentTable.rows.map((row) => ({
        ...row,
        ...(documentEdits[row.rowNumber] ?? {}),
      })),
    [documentEdits, documentTable],
  )

  const documentChanges = useMemo(
    () =>
      documentTable.rows.flatMap((originalRow) =>
        DOCUMENT_EDIT_FIELDS.flatMap(({ key }) => {
          const editedValue = documentEdits[originalRow.rowNumber]?.[key]
          if (editedValue === undefined || editedValue === originalRow[key]) return []

          const columnIndex = documentTable.columns[key]
          if (!Number.isInteger(columnIndex) || columnIndex < 0) return []

          return {
            id: `document-${originalRow.rowNumber}-${key}`,
            address: `${getColumnLetter(columnIndex)}${originalRow.rowNumber}`,
            value: editedValue,
          }
        }),
      ),
    [documentEdits, documentTable],
  )

  const viewerSheet = viewerWorkbook && viewerSheetName
    ? viewerWorkbook.Sheets[viewerSheetName]
    : null
  const viewerSheetView = useMemo(
    () => buildSheetView(viewerSheet),
    [viewerSheet],
  )

  useEffect(() => {
    const restoreDraft = async () => {
      const rawDraft = localStorage.getItem(WORK_DRAFT_KEY)
      if (!rawDraft) return

      try {
        const draft = JSON.parse(rawDraft)
        const buffer = await dataUrlToArrayBuffer(draft.fileDataUrl)
        const restoredWorkbook = XLSX.read(buffer, { type: 'array', cellDates: true })

        setWorkbook(restoredWorkbook)
        setSourceBuffer(buffer)
        setFileName(draft.fileName ?? '')
        setSheetName(draft.sheetName ?? restoredWorkbook.SheetNames[0] ?? '')
        setNameColumn(draft.nameColumn ?? '')
        setManualPeople(draft.manualPeople ?? [])
        setDocumentEdits(draft.documentEdits ?? {})
        setRulesByTarget(draft.rulesByTarget ?? {
          actNumber: [createRule('actNumber')],
          erdrExtract: [createRule('erdrExtract')],
        })
        setActiveTargetType(draft.activeTargetType ?? 'actNumber')
        setActiveRuleId(draft.activeRuleId ?? '')
        setStatus(`Восстановлен файл: ${draft.fileName}`)
      } catch {
        localStorage.removeItem(WORK_DRAFT_KEY)
      }
    }

    restoreDraft()

  }, [])

  useEffect(() => {
    if (!sourceBuffer || !fileName) return

    const saveDraft = async () => {
      try {
        const fileDataUrl = await arrayBufferToDataUrl(sourceBuffer)
        localStorage.setItem(
          WORK_DRAFT_KEY,
          JSON.stringify({
            fileDataUrl,
            fileName,
            sheetName,
            nameColumn,
            manualPeople,
            documentEdits,
            rulesByTarget,
            activeTargetType,
            activeRuleId,
            savedAt: new Date().toISOString(),
          }),
        )
      } catch {
        setStatus('Не удалось сохранить черновик в браузере')
      }
    }

    saveDraft()
  }, [
    sourceBuffer,
    fileName,
    sheetName,
    nameColumn,
    manualPeople,
    documentEdits,
    rulesByTarget,
    activeTargetType,
    activeRuleId,
  ])

  useEffect(() => {
    if (!fileName || !sheetName) return

    localStorage.setItem(
      DOCUMENT_EDITS_DRAFT_KEY,
      JSON.stringify({
        fileName,
        sheetName,
        documentEdits,
        savedAt: new Date().toISOString(),
      }),
    )
  }, [documentEdits, fileName, sheetName])

  const usedColumns = useMemo(() => {
    const maxColumns = rows.reduce((max, row) => Math.max(max, row.length), 0)
    return Array.from({ length: maxColumns }, (_, index) => {
      const samples = rows
        .slice(0, 6)
        .map((row) => row[index])
        .filter((cell) => String(cell).trim())
        .slice(0, 3)

      return {
        index,
        letter: getColumnLetter(index),
        label: `${getColumnLetter(index)}${samples.length ? ` - ${samples.join(', ')}` : ''}`,
      }
    })
  }, [rows])

  const detectedPeople = useMemo(() => {
    if (!nameColumn) return []
    const colIndex = XLSX.utils.decode_col(nameColumn)
    const found = new Map()

    rows.forEach((row, index) => {
      const raw = row[colIndex]
      const fullName = String(raw ?? '').trim()
      if (!isLikelyName(fullName)) return

      const surname = fullName.split(/\s+/)[0]
      const key = normalize(surname)
      const targetRows = findTargetRowsForPerson(rows, index)
      if (!found.has(key)) {
        found.set(key, { surname, fullName, rows: [], occurrences: [] })
      }
      found.get(key).rows.push(index + 1)
      found.get(key).occurrences.push({
        nameRow: index + 1,
        targetRows,
      })
    })

    return [...found.values()].sort((a, b) =>
      a.surname.localeCompare(b.surname, 'ru'),
    )
  }, [nameColumn, rows])

  const people = useMemo(() => {
    const merged = new Map(
      detectedPeople.map((person) => [
        normalize(person.surname),
        { ...person, manual: false },
      ]),
    )

    manualPeople.forEach((person) => {
      const key = normalize(person.surname)
      const existing = merged.get(key)
      if (existing) {
        merged.set(key, {
          ...existing,
          rows: [...new Set([...existing.rows, ...person.rows])].sort((a, b) => a - b),
          occurrences: [...(existing.occurrences ?? []), ...(person.occurrences ?? [])],
          manual: existing.manual || person.manual,
        })
        return
      }

      merged.set(key, person)
    })

    return [...merged.values()].sort((a, b) =>
      a.surname.localeCompare(b.surname, 'ru'),
    )
  }, [detectedPeople, manualPeople])

  const filteredPeople = useMemo(() => {
    const query = normalize(searchTerm)
    if (!query) return people
    return people.filter(
      (person) =>
        normalize(person.surname).includes(query) ||
        normalize(person.fullName).includes(query),
    )
  }, [people, searchTerm])

  const visibleRules = rulesByTarget[activeTargetType] ?? []
  const activeRule =
    visibleRules.find((rule) => rule.id === activeRuleId) ?? visibleRules[0] ?? null
  const selectedSet = useMemo(
    () => new Set(activeRule?.personKeys ?? []),
    [activeRule],
  )
  const activeSelectedCount = activeRule?.personKeys?.length ?? 0
  const cleanRules = useMemo(
    () =>
      Object.values(rulesByTarget)
        .flat()
        .map((rule) => {
          const day = Number.parseInt(rule.day, 10)
          const column = getColumnByDay(rule.day)
          return {
            ...rule,
            day: Number.isInteger(day) && day >= 1 && day <= 30 ? day : null,
            column,
            personKeys: rule.personKeys ?? [],
            targetLabel: TARGET_TYPES[rule.targetType]?.label ?? 'Номер акта',
          }
        })
        .filter((rule) => rule.column),
    [rulesByTarget],
  )

  const previewChanges = useMemo(() => {
    const peopleBySurname = new Map(
      people.map((person) => [normalize(person.surname), person]),
    )

    return cleanRules.flatMap((rule) =>
      (rule.personKeys ?? []).flatMap((personKey) => {
        const person = peopleBySurname.get(personKey)
        if (!person) return []

        return (person.occurrences ?? []).map((occurrence) => {
          const targetRow = occurrence.targetRows?.[rule.targetType]
          const address = targetRow ? `${rule.column}${targetRow}` : ''

          return {
            id: `${personKey}-${rule.id}-${occurrence.nameRow}-${address || 'missing'}`,
            surname: person.surname,
            targetLabel: rule.targetLabel,
            day: rule.day,
            row: targetRow || '',
            column: rule.column,
            address,
            value: rule.value,
            ready: Boolean(address),
          }
        })
      }),
    )
  }, [cleanRules, people])

  const handleLogin = (event) => {
    event.preventDefault()
    if (login === DEMO_LOGIN && password === DEMO_PASSWORD) {
      localStorage.setItem('excel-tool-auth', 'yes')
      setIsAuthorized(true)
      setAuthError('')
      return
    }

    setAuthError('Неверный логин или пароль')
  }

  const handleLogout = () => {
    localStorage.removeItem('excel-tool-auth')
    setIsAuthorized(false)
    setLogin('')
    setPassword('')
  }

  const handleUpload = async (event) => {
    const file = event.target.files?.[0]
    if (!file) return

    const buffer = await file.arrayBuffer()
    await openHelperWorkbook(buffer, file.name, true)
  }

  const openHelperWorkbook = async (buffer, name, rememberFile = false) => {
    const loadedWorkbook = XLSX.read(buffer, { type: 'array', cellDates: true })
    const firstSheet = loadedWorkbook.SheetNames[0] ?? ''

    setWorkbook(loadedWorkbook)
    setSourceBuffer(buffer)
    setFileName(name)
    setSheetName(firstSheet)
    setRulesByTarget({
      actNumber: [createRule('actNumber')],
      erdrExtract: [createRule('erdrExtract')],
    })
    setActiveRuleId('')
    setManualPeople([])
    try {
      const documentDraft = JSON.parse(
        localStorage.getItem(DOCUMENT_EDITS_DRAFT_KEY) ?? '{}',
      )
      setDocumentEdits(
        documentDraft.fileName === name && documentDraft.sheetName === firstSheet
          ? documentDraft.documentEdits ?? {}
          : {},
      )
    } catch {
      localStorage.removeItem(DOCUMENT_EDITS_DRAFT_KEY)
      setDocumentEdits({})
    }
    setSearchTerm('')
    setStatus(
      name.toLowerCase().includes('updated')
        ? `Файл загружен: ${name}. Лучше взять исходный файл без updated в имени.`
        : `Файл загружен: ${name}`,
    )

    const firstRows = XLSX.utils.sheet_to_json(loadedWorkbook.Sheets[firstSheet], {
      header: 1,
      defval: '',
    })
    const columnScores = []
    const maxColumns = firstRows.reduce((max, row) => Math.max(max, row.length), 0)
    for (let col = 0; col < maxColumns; col += 1) {
      const score = firstRows
        .slice(0, 80)
        .filter((row) => isLikelyName(row[col])).length
      columnScores.push({ col, score })
    }
    const bestColumn = columnScores.sort((a, b) => b.score - a.score)[0]
    setNameColumn(bestColumn?.score ? getColumnLetter(bestColumn.col) : '')

    if (rememberFile) {
      const fileDataUrl = await arrayBufferToDataUrl(buffer)
      const nextRecentFiles = [
      {
          id: `${name}-${buffer.byteLength}`,
          name,
          savedAt: new Date().toISOString(),
          fileDataUrl,
        },
        ...recentHelperFiles.filter((recentFile) => recentFile.name !== name),
      ].slice(0, 3)
      setRecentHelperFiles(nextRecentFiles)
      setActiveHelperRecentId(nextRecentFiles[0].id)
      localStorage.setItem(RECENT_HELPER_FILES_KEY, JSON.stringify(nextRecentFiles))
    }
  }

  const openRecentHelperFile = async (recentFile) => {
    const buffer = await dataUrlToArrayBuffer(recentFile.fileDataUrl)
    await openHelperWorkbook(buffer, recentFile.name)
    setActiveHelperRecentId(recentFile.id)
  }

  const clearHelperFile = () => {
    setWorkbook(null)
    setSourceBuffer(null)
    setFileName('')
    setSheetName('')
    setNameColumn('')
    setSearchTerm('')
    setManualPeople([])
    setDocumentEdits({})
    setRulesByTarget({
      actNumber: [createRule('actNumber')],
      erdrExtract: [createRule('erdrExtract')],
    })
    setActiveRuleId('')
    setActiveHelperRecentId('')
    setStatus('Файл убран с сайта')
    localStorage.removeItem(WORK_DRAFT_KEY)
    localStorage.removeItem(DOCUMENT_EDITS_DRAFT_KEY)
    if (fileInputRef.current) {
      fileInputRef.current.value = ''
    }
  }

  const removeRecentHelperFile = (fileId) => {
    const nextRecentFiles = recentHelperFiles.filter(
      (recentFile) => recentFile.id !== fileId,
    )
    setRecentHelperFiles(nextRecentFiles)
    localStorage.setItem(RECENT_HELPER_FILES_KEY, JSON.stringify(nextRecentFiles))

    if (activeHelperRecentId === fileId) {
      setActiveHelperRecentId('')
    }
  }

  const handleViewerUpload = async (event) => {
    const file = event.target.files?.[0]
    if (!file) return

    const buffer = await file.arrayBuffer()
    const loadedWorkbook = XLSX.read(buffer, { type: 'array', cellDates: true })
    const firstSheet = loadedWorkbook.SheetNames[0] ?? ''

    setViewerWorkbook(loadedWorkbook)
    setViewerFileName(file.name)
    setViewerSheetName(firstSheet)
    setViewerZoom(1)

    const fileDataUrl = await arrayBufferToDataUrl(buffer)
    const nextRecentFiles = [
      {
        id: `${file.name}-${file.size}-${file.lastModified}`,
        name: file.name,
        savedAt: new Date().toISOString(),
        fileDataUrl,
      },
      ...recentViewerFiles.filter((recentFile) => recentFile.name !== file.name),
    ].slice(0, 3)
    setRecentViewerFiles(nextRecentFiles)
    setActiveViewerRecentId(nextRecentFiles[0].id)
    localStorage.setItem(RECENT_VIEWER_FILES_KEY, JSON.stringify(nextRecentFiles))
  }

  const openRecentViewerFile = async (recentFile) => {
    const buffer = await dataUrlToArrayBuffer(recentFile.fileDataUrl)
    const loadedWorkbook = XLSX.read(buffer, { type: 'array', cellDates: true })
    const firstSheet = loadedWorkbook.SheetNames[0] ?? ''

    setViewerWorkbook(loadedWorkbook)
    setViewerFileName(recentFile.name)
    setViewerSheetName(firstSheet)
    setViewerZoom(1)
    setActiveViewerRecentId(recentFile.id)
  }

  const clearViewerFile = () => {
    setViewerWorkbook(null)
    setViewerFileName('')
    setViewerSheetName('')
    setViewerZoom(1)
    setActiveViewerRecentId('')
    if (viewerFileInputRef.current) {
      viewerFileInputRef.current.value = ''
    }
  }

  const removeRecentViewerFile = (fileId) => {
    const nextRecentFiles = recentViewerFiles.filter(
      (recentFile) => recentFile.id !== fileId,
    )
    setRecentViewerFiles(nextRecentFiles)
    localStorage.setItem(RECENT_VIEWER_FILES_KEY, JSON.stringify(nextRecentFiles))

    if (activeViewerRecentId === fileId) {
      setActiveViewerRecentId('')
    }
  }

  const toggleName = (surname) => {
    const personKey = normalize(surname)
    const targetRuleId = activeRule?.id
    if (!targetRuleId) return

    setRulesByTarget((current) => ({
      ...current,
      [activeTargetType]: current[activeTargetType].map((rule) => {
        if (rule.id !== targetRuleId) return rule

        const currentKeys = rule.personKeys ?? []
        return {
          ...rule,
          personKeys: currentKeys.includes(personKey)
            ? currentKeys.filter((key) => key !== personKey)
            : [...currentKeys, personKey],
        }
      }),
    }))
  }

  const addManualPerson = (event) => {
    event.preventDefault()
    const fullName = manualName.trim()
    if (!fullName) return

    const surname = fullName.split(/\s+/)[0]
    const requestedRow = Number.parseInt(manualRow, 10)
    const matchingRows = nameColumn
      ? rows
          .map((row, index) => ({
            rowNumber: index + 1,
            value: String(row[XLSX.utils.decode_col(nameColumn)] ?? '').trim(),
          }))
          .filter((row) => normalize(row.value.split(/\s+/)[0]) === normalize(surname))
          .map((row) => row.rowNumber)
      : []
    const personRows = Number.isInteger(requestedRow) && requestedRow > 0
      ? [requestedRow]
      : matchingRows
    const occurrences = personRows.map((rowNumber) => ({
      nameRow: rowNumber,
      targetRows: findTargetRowsForPerson(rows, rowNumber - 1),
    }))

    setManualPeople((current) => {
      const key = normalize(surname)
      const existing = current.find((person) => normalize(person.surname) === key)
      if (existing) {
        return current.map((person) =>
          normalize(person.surname) === key
            ? {
                ...person,
                fullName,
                rows: [...new Set([...person.rows, ...personRows])].sort((a, b) => a - b),
                occurrences: [...(person.occurrences ?? []), ...occurrences],
              }
            : person,
        )
      }

      return [
        ...current,
        {
          surname,
          fullName,
          rows: personRows,
          occurrences,
          manual: true,
        },
      ]
    })
    const personKey = normalize(surname)
    const targetRuleId = activeRule?.id
    if (targetRuleId) {
      setRulesByTarget((current) => ({
        ...current,
        [activeTargetType]: current[activeTargetType].map((rule) =>
          rule.id === targetRuleId
            ? { ...rule, personKeys: [...new Set([...(rule.personKeys ?? []), personKey])] }
            : rule,
        ),
      }))
    }
    setManualName('')
    setManualRow('')
  }

  const removeManualPerson = (surname) => {
    const key = normalize(surname)
    setManualPeople((current) =>
      current.filter((person) => normalize(person.surname) !== key),
    )
    setRulesByTarget((current) =>
      Object.fromEntries(
        Object.entries(current).map(([targetType, rules]) => [
          targetType,
          rules.map((rule) => ({
            ...rule,
            personKeys: (rule.personKeys ?? []).filter((personKey) => personKey !== key),
          })),
        ]),
      ),
    )
  }

  const toggleAllVisible = () => {
    const visibleKeys = filteredPeople.map((person) => normalize(person.surname))
    const allSelected = visibleKeys.every((personKey) => selectedSet.has(personKey))
    const targetRuleId = activeRule?.id
    if (!targetRuleId) return

    setRulesByTarget((current) => ({
      ...current,
      [activeTargetType]: current[activeTargetType].map((rule) => {
        if (rule.id !== targetRuleId) return rule
        const currentKeys = rule.personKeys ?? []
        return {
          ...rule,
          personKeys: allSelected
            ? currentKeys.filter((personKey) => !visibleKeys.includes(personKey))
            : [...new Set([...currentKeys, ...visibleKeys])],
        }
      }),
    }))
  }

  const updateRule = (id, field, value) => {
    setRulesByTarget((current) => ({
      ...current,
      [activeTargetType]: current[activeTargetType].map((rule) =>
        rule.id === id ? { ...rule, [field]: value } : rule,
      ),
    }))
  }

  const updateDocumentRow = (rowNumber, field, value) => {
    setDocumentEdits((current) => ({
      ...current,
      [rowNumber]: {
        ...(current[rowNumber] ?? {}),
        [field]: value,
      },
    }))
  }

  const addRule = () => {
    const newRule = createRule(activeTargetType)
    setRulesByTarget((current) => ({
      ...current,
      [activeTargetType]: [
        ...(current[activeTargetType] ?? []),
        newRule,
      ],
    }))
    setActiveRuleId(newRule.id)
  }

  const removeRule = (id) => {
    setRulesByTarget((current) => {
      const currentRules = current[activeTargetType] ?? []
      if (currentRules.length === 1) return current

      return {
        ...current,
        [activeTargetType]: currentRules.filter((rule) => rule.id !== id),
      }
    })
  }

  const handleTargetChange = (targetType) => {
    setActiveTargetType(targetType)
  }

  const buildUpdatedWorkbook = async () => {
    if (!workbook || !activeSheet || !sourceBuffer) {
      return { error: 'Сначала загрузите Excel-файл' }
    }

    if (!previewChanges.length && !documentChanges.length) {
      return { error: 'Выберите фамилии или измените нижнюю таблицу документов' }
    }

    if (previewChanges.length && !cleanRules.length) {
      return { error: 'Укажите хотя бы одну дату для заполнения' }
    }

    const zip = await JSZip.loadAsync(sourceBuffer.slice(0))
    const sheetXmlPath = await getSheetXmlPath(zip, sheetName)
    const sheetFile = sheetXmlPath ? zip.file(sheetXmlPath) : null
    if (!sheetFile) {
      return { error: 'Не удалось найти лист в структуре .xlsx' }
    }

    let sheetXml = await sheetFile.async('string')
    const sharedStringsFile = zip.file('xl/sharedStrings.xml')
    if (!sharedStringsFile) {
      return { error: 'В файле не найдена таблица строк sharedStrings.xml' }
    }

    let sharedStringsXml = await sharedStringsFile.async('string')
    let changedCells = 0

    const writeChange = (change) => {
      const valuePayload = getCellXml(change.value)
      if (valuePayload.valueXml) {
        sheetXml = replaceCellValue(sheetXml, change.address, valuePayload)
      } else {
        const sharedString = appendSharedString(sharedStringsXml, change.value)
        sharedStringsXml = sharedString.xml
        sheetXml = replaceCellValue(sheetXml, change.address, {
          ...valuePayload,
          sharedStringIndex: sharedString.index,
        })
      }
      changedCells += 1
    }

    previewChanges.forEach((change) => {
      if (!change.ready) return
      writeChange(change)
    })

    documentChanges.forEach((change) => {
      writeChange(change)
    })

    if (!changedCells) {
      return { error: 'Нет изменений для записи в Excel' }
    }

    zip.file(sheetXmlPath, sheetXml)
    zip.file('xl/sharedStrings.xml', sharedStringsXml)
    const safeName = fileName.replace(/\.[^.]+$/, '') || 'table'
    const blob = await zip.generateAsync({
      type: 'blob',
      compression: 'DEFLATE',
    })

    return {
      blob,
      changedCells,
      outputName: `${safeName}_updated.xlsx`,
    }
  }

  const saveCurrentWork = async () => {
    const result = await buildUpdatedWorkbook()
    if (result.error) {
      setStatus(result.error)
      return
    }

    if (CLOUD_SAVE_URL) {
      const formData = new FormData()
      formData.append('file', result.blob, result.outputName)
      formData.append('fileName', result.outputName)
      formData.append('appVersion', APP_VERSION)
      formData.append('savedAt', new Date().toISOString())

      const response = await fetch(CLOUD_SAVE_URL, {
        method: 'POST',
        body: formData,
      })

      if (!response.ok) {
        setStatus('Не удалось сохранить Excel в облаке')
        return
      }

      setStatus(`Excel сохранен в облаке: ${result.outputName}`)
      return
    }

    const reader = new FileReader()
    reader.onload = () => {
      localStorage.setItem(
        'police-helper-xlsx-save',
        JSON.stringify({
          fileName: result.outputName,
          savedAt: new Date().toISOString(),
          dataUrl: reader.result,
        }),
      )
      setStatus(`Excel сохранен в браузере: ${result.outputName}`)
    }
    reader.readAsDataURL(result.blob)
  }

  const downloadWorkbook = async () => {
    const result = await buildUpdatedWorkbook()
    if (result.error) {
      setStatus(result.error)
      return
    }

    saveAs(
      result.blob,
      result.outputName,
    )
    setStatus(`Готово: изменено ячеек ${result.changedCells}`)
  }

  if (!isAuthorized) {
    return (
      <main className="login-page">
        <form className="login-panel" onSubmit={handleLogin}>
          <div className="brand-mark">
            <FileSpreadsheet size={34} aria-hidden="true" />
          </div>
          <div>
            <h1>Excel кабинет</h1>
          </div>

          <label>
            <span>Логин</span>
            <div className="field">
              <UserRound size={18} aria-hidden="true" />
              <input
                value={login}
                onChange={(event) => setLogin(event.target.value)}
                placeholder="admin"
                autoComplete="username"
              />
            </div>
          </label>

          <label>
            <span>Пароль</span>
            <div className="field">
              <KeyRound size={18} aria-hidden="true" />
              <input
                value={password}
                onChange={(event) => setPassword(event.target.value)}
                placeholder="12345"
                type="password"
                autoComplete="current-password"
              />
            </div>
          </label>

          {authError && <p className="error">{authError}</p>}

          <button className="primary-button" type="submit">
            <Check size={18} aria-hidden="true" />
            Войти
          </button>
        </form>
      </main>
    )
  }

  return (
    <main className="app-shell">
      <header className="topbar">
        <div>
          <p className="eyebrow">Excel кабинет</p>
          <h1 className="app-title">Police Helper</h1>
        </div>
        <div className="top-actions">
          <nav className="page-tabs" aria-label="Навигация">
            <button
              className={currentPage === 'helper' ? 'page-tab active' : 'page-tab'}
              type="button"
              onClick={() => setCurrentPage('helper')}
            >
              Police Helper
            </button>
            <button
              className={currentPage === 'excel' ? 'page-tab active' : 'page-tab'}
              type="button"
              onClick={() => setCurrentPage('excel')}
            >
              Excel
            </button>
          </nav>
          <button className="icon-button logout-button" type="button" onClick={handleLogout} title="Выйти">
            <LogOut size={20} aria-hidden="true" />
          </button>
        </div>
      </header>

      {currentPage === 'excel' ? (
        <section className="excel-viewer-page">
          <div className="panel excel-viewer-panel">
            <div className="panel-heading">
              <div>
                <h2>Excel</h2>
                <p>Загрузите файл и просматривайте таблицу прямо на сайте.</p>
              </div>
              <button
                className="secondary-button"
                type="button"
                onClick={() => viewerFileInputRef.current?.click()}
              >
                <Upload size={18} aria-hidden="true" />
                Загрузить Excel
              </button>
            </div>

            {recentViewerFiles.length > 0 && (
              <div className="recent-files">
                {recentViewerFiles.map((recentFile) => (
                  <div
                    className={
                      activeViewerRecentId === recentFile.id
                        ? 'recent-file active'
                        : 'recent-file'
                    }
                    key={recentFile.id}
                    title={recentFile.name}
                  >
                    <button
                      className="recent-file-open"
                      type="button"
                      onClick={() => openRecentViewerFile(recentFile)}
                    >
                      <FileSpreadsheet size={20} aria-hidden="true" />
                      <span>{recentFile.name}</span>
                    </button>
                    <button
                      className="recent-file-delete"
                      type="button"
                      onClick={() => removeRecentViewerFile(recentFile.id)}
                      title="Удалить из последних"
                    >
                      <Trash2 size={14} aria-hidden="true" />
                    </button>
                  </div>
                ))}
              </div>
            )}

            <input
              ref={viewerFileInputRef}
              className="hidden-input"
              type="file"
              accept=".xlsx,.xls"
              onChange={handleViewerUpload}
            />

            <div className={viewerFileName ? 'excel-viewer-controls has-file' : 'excel-viewer-controls'}>
              {viewerFileName && (
                <div className="file-chip">
                  <FileSpreadsheet size={18} aria-hidden="true" />
                  <span>{viewerFileName}</span>
                </div>
              )}

              {viewerFileName && (
                <button
                  className="icon-button danger"
                  type="button"
                  onClick={clearViewerFile}
                  title="Убрать загруженный файл"
                >
                  <Trash2 size={18} aria-hidden="true" />
                </button>
              )}

              <label>
                <span>Лист</span>
                <select
                  value={viewerSheetName}
                  onChange={(event) => setViewerSheetName(event.target.value)}
                  disabled={!viewerWorkbook}
                >
                  {(viewerWorkbook?.SheetNames ?? []).map((name) => (
                    <option key={name} value={name}>
                      {name}
                    </option>
                  ))}
                </select>
              </label>

              <div className="zoom-controls">
                <button
                  className="icon-button"
                  type="button"
                  onClick={() => setViewerZoom((zoom) => Math.max(0.4, zoom - 0.1))}
                  title="Уменьшить"
                >
                  -
                </button>
                <span>{Math.round(viewerZoom * 100)}%</span>
                <button
                  className="icon-button"
                  type="button"
                  onClick={() => setViewerZoom((zoom) => Math.min(2.5, zoom + 0.1))}
                  title="Увеличить"
                >
                  +
                </button>
              </div>
            </div>

            <div className="excel-grid-wrap">
              {viewerSheetView.cells.length ? (
                <table
                  className="excel-grid"
                  style={{ transform: `scale(${viewerZoom})` }}
                >
                  <colgroup>
                    {viewerSheetView.columnWidths.map((width, index) => (
                      <col key={index} style={{ width: `${width}px` }} />
                    ))}
                  </colgroup>
                  <tbody>
                    {viewerSheetView.cells.map((row, rowIndex) => (
                      <tr
                        key={`row-${rowIndex}`}
                        style={row.height ? { height: `${row.height}px` } : undefined}
                      >
                        {row.cells.map((cell) => (
                          <td
                            colSpan={cell.colSpan}
                            key={cell.address}
                            rowSpan={cell.rowSpan}
                            title={String(cell.value)}
                          >
                            {String(cell.value)}
                          </td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              ) : (
                <div className="empty-state">Загрузите Excel-файл для просмотра.</div>
              )}
            </div>
          </div>
        </section>
      ) : (
      <section className="workspace">
        <aside className="panel upload-panel">
          <div>
            <h2>Файл</h2>
          </div>

          <input
            ref={fileInputRef}
            className="hidden-input"
            type="file"
            accept=".xlsx,.xls"
            onChange={handleUpload}
          />
          <button
            className="secondary-button"
            type="button"
            onClick={() => fileInputRef.current?.click()}
          >
            <Upload size={18} aria-hidden="true" />
            Загрузить Excel
          </button>

          {recentHelperFiles.length > 0 && (
            <div className="recent-files compact">
              {recentHelperFiles.map((recentFile) => (
                <div
                  className={
                    activeHelperRecentId === recentFile.id
                      ? 'recent-file active'
                      : 'recent-file'
                  }
                  key={recentFile.id}
                  title={recentFile.name}
                >
                  <button
                    className="recent-file-open"
                    type="button"
                    onClick={() => openRecentHelperFile(recentFile)}
                  >
                    <FileSpreadsheet size={20} aria-hidden="true" />
                    <span>{recentFile.name}</span>
                  </button>
                  <button
                    className="recent-file-delete"
                    type="button"
                    onClick={() => removeRecentHelperFile(recentFile.id)}
                    title="Удалить из последних"
                  >
                    <Trash2 size={14} aria-hidden="true" />
                  </button>
                </div>
              ))}
            </div>
          )}

          {fileName && (
            <div className="file-chip-row">
              <div className="file-chip">
                <FileSpreadsheet size={18} aria-hidden="true" />
                <span>{fileName}</span>
              </div>
              <button
                className="icon-button danger"
                type="button"
                onClick={clearHelperFile}
                title="Убрать загруженный файл"
              >
                <Trash2 size={18} aria-hidden="true" />
              </button>
            </div>
          )}

          <label>
            <span>Лист</span>
            <select
              value={sheetName}
              onChange={(event) => {
                setSheetName(event.target.value)
                setRulesByTarget({
                  actNumber: [createRule('actNumber')],
                  erdrExtract: [createRule('erdrExtract')],
                })
                setActiveRuleId('')
              }}
              disabled={!workbook}
            >
              {(workbook?.SheetNames ?? []).map((name) => (
                <option key={name} value={name}>
                  {name}
                </option>
              ))}
            </select>
          </label>

          <label>
            <span>Колонка с фамилиями</span>
            <select
              value={nameColumn}
              onChange={(event) => {
                setNameColumn(event.target.value)
                setRulesByTarget({
                  actNumber: [createRule('actNumber')],
                  erdrExtract: [createRule('erdrExtract')],
                })
                setActiveRuleId('')
              }}
              disabled={!usedColumns.length}
            >
              <option value="">Выберите колонку</option>
              {usedColumns.map((column) => (
                <option key={column.letter} value={column.letter}>
                  {column.label}
                </option>
              ))}
            </select>
          </label>

          <div className="status-box">{status || 'Ожидаю загрузку таблицы'}</div>
        </aside>

        <section className="panel people-panel">
          <div className="panel-heading">
            <div>
              <h2>Фамилии</h2>
              <p>Найдено: {people.length}. В активном поле: {activeSelectedCount}.</p>
            </div>
            <button
              className="ghost-button"
              type="button"
              onClick={toggleAllVisible}
              disabled={!filteredPeople.length}
            >
              {filteredPeople.every((person) => selectedSet.has(normalize(person.surname)))
                ? 'Снять'
                : 'Выбрать'}
            </button>
          </div>

          <div className="search-field">
            <Search size={18} aria-hidden="true" />
            <input
              value={searchTerm}
              onChange={(event) => setSearchTerm(event.target.value)}
              placeholder="Поиск фамилии"
            />
          </div>

          <div className="people-list">
            {filteredPeople.map((person) => (
              <div className="person-row" key={person.surname}>
                <input
                  type="checkbox"
                  checked={selectedSet.has(normalize(person.surname))}
                  onChange={() => toggleName(person.surname)}
                />
                <span>
                  <strong>{person.surname}</strong>
                  <small>
                    {person.fullName} · {person.rows.length ? `строки ${person.rows.join(', ')}` : 'строка не указана'}
                  </small>
                </span>
                {person.manual && (
                  <button
                    className="icon-button danger"
                    type="button"
                    onClick={() => removeManualPerson(person.surname)}
                    title="Удалить ручную фамилию"
                  >
                    <Trash2 size={18} aria-hidden="true" />
                  </button>
                )}
              </div>
            ))}

            {!filteredPeople.length && (
              <div className="empty-state">
                Загрузите файл и выберите колонку, где находятся фамилии.
              </div>
            )}
          </div>
        </section>

        <aside className="panel rules-panel">
          <div className="panel-heading">
            <div>
              <h2>Что вставить</h2>
            </div>
            <button className="icon-button" type="button" onClick={addRule} title="Добавить поле">
              <Plus size={20} aria-hidden="true" />
            </button>
          </div>

          <div className="target-tabs" role="tablist" aria-label="Тип данных">
            {Object.entries(TARGET_TYPES).map(([key, target]) => (
              <button
                className={activeTargetType === key ? 'target-tab active' : 'target-tab'}
                key={key}
                type="button"
                onClick={() => handleTargetChange(key)}
              >
                {target.label}
              </button>
            ))}
          </div>

          <div className="rules-list">
            {visibleRules.map((rule) => (
              <div
                className={activeRule?.id === rule.id ? 'rule-row active' : 'rule-row'}
                key={rule.id}
                onClick={() => setActiveRuleId(rule.id)}
              >
                <label>
                  <span>Дата</span>
                  <select
                    value={rule.day}
                    onChange={(event) => updateRule(rule.id, 'day', event.target.value)}
                  >
                    <option value="">Дата</option>
                    {Array.from({ length: 30 }, (_, index) => {
                      const day = String(index + 1).padStart(2, '0')
                      return (
                        <option key={day} value={String(index + 1)}>
                          {day}
                        </option>
                      )
                    })}
                  </select>
                </label>
                <label>
                  <span>Значение</span>
                  <input
                    value={rule.value}
                    onChange={(event) => updateRule(rule.id, 'value', event.target.value)}
                    placeholder="Текст или число"
                  />
                </label>
                <button
                  className="icon-button danger"
                  type="button"
                  onClick={() => removeRule(rule.id)}
                  title="Удалить поле"
                >
                  <Trash2 size={18} aria-hidden="true" />
                </button>
              </div>
            ))}
          </div>

          <button className="secondary-button download-button" type="button" onClick={saveCurrentWork}>
            <Save size={18} aria-hidden="true" />
            Сохранить
          </button>

          <button className="primary-button download-button" type="button" onClick={downloadWorkbook}>
            <Download size={18} aria-hidden="true" />
            Скачать Excel
          </button>
        </aside>

        <section className="panel documents-panel">
          <div className="panel-heading">
            <div>
              <h2>Реквізити документів</h2>
              <p>Строк: {documentRows.length}. Изменений: {documentChanges.length}.</p>
            </div>
          </div>

          <div className="documents-table-wrap">
            <table className="documents-table">
              <thead>
                <tr>
                  <th>Вид документа</th>
                  {DOCUMENT_EDIT_FIELDS.map((field) => (
                    <th key={field.key}>{field.label}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {documentRows.map((row) => (
                  <tr key={row.rowNumber}>
                    <td>{row.type}</td>
                    {DOCUMENT_EDIT_FIELDS.map((field) => {
                      const isLockedAgency =
                        field.key === 'agency' && row.type === 'Акт огляду'

                      return (
                        <td key={field.key}>
                          {isLockedAgency ? (
                            <span className="document-static-cell">{row[field.key]}</span>
                          ) : field.key === 'agency' && row.type === 'ЄО' ? (
                            <select
                              value={row[field.key]}
                              onChange={(event) =>
                                updateDocumentRow(row.rowNumber, field.key, event.target.value)
                              }
                            >
                              <option value="">Выберите орган</option>
                              {EO_AGENCY_OPTIONS.map((agency) => (
                                <option key={agency} value={agency}>
                                  {agency}
                                </option>
                              ))}
                            </select>
                          ) : field.key === 'signer' && row.type === 'Акт огляду' ? (
                            <select
                              value={row[field.key]}
                              onChange={(event) =>
                                updateDocumentRow(row.rowNumber, field.key, event.target.value)
                              }
                            >
                              <option value="">Выберите подписанта</option>
                              {ACT_SIGNER_OPTIONS.map((signer) => (
                                <option key={signer} value={signer}>
                                  {signer}
                                </option>
                              ))}
                            </select>
                          ) : field.key === 'signer' && row.type === 'ЄО' ? (
                            <select
                              value={row[field.key]}
                              onChange={(event) =>
                                updateDocumentRow(row.rowNumber, field.key, event.target.value)
                              }
                            >
                              <option value="">Выберите подписанта</option>
                              {EO_SIGNER_OPTIONS.map((signer) => (
                                <option key={signer} value={signer}>
                                  {signer}
                                </option>
                              ))}
                            </select>
                          ) : field.key === 'date' && row.type !== 'ЄО' ? (
                          <select
                            value={row[field.key]}
                            onChange={(event) =>
                              updateDocumentRow(row.rowNumber, field.key, event.target.value)
                            }
                          >
                            <option value="">Дата</option>
                            {MAY_2026_DATES.map((date) => (
                              <option key={date} value={date}>
                                {date}
                              </option>
                            ))}
                          </select>
                        ) : (
                          <textarea
                            value={row[field.key]}
                            onChange={(event) =>
                              updateDocumentRow(row.rowNumber, field.key, event.target.value)
                            }
                            rows={field.key === 'number' || field.key === 'date' ? 1 : 2}
                          />
                        )}
                        </td>
                      )
                    })}
                  </tr>
                ))}
              </tbody>
            </table>

            {!documentRows.length && (
              <div className="empty-state">
                Нижняя таблица документов появится после загрузки подходящего Excel-файла.
              </div>
            )}
          </div>
        </section>

        <section className="panel preview-panel">
          <div className="panel-heading">
            <div>
              <h2>Предпросмотр</h2>
              <p>Изменений: {previewChanges.filter((change) => change.ready).length}</p>
            </div>
          </div>

          <div className="preview-table-wrap">
            <table className="preview-table">
              <thead>
                <tr>
                  <th>Фамилия</th>
                  <th>Что</th>
                  <th>Дата</th>
                  <th>Ячейка</th>
                  <th>Строка</th>
                  <th>Колонка</th>
                  <th>Значение</th>
                </tr>
              </thead>
              <tbody>
                {previewChanges.slice(0, 80).map((change) => (
                  <tr className={change.ready ? '' : 'muted-row'} key={change.id}>
                    <td>{change.surname}</td>
                    <td>{change.targetLabel}</td>
                    <td>{String(change.day).padStart(2, '0')}</td>
                    <td>{change.address || 'нет строки'}</td>
                    <td>{change.row || '-'}</td>
                    <td>{change.column}</td>
                    <td>{change.value}</td>
                  </tr>
                ))}
              </tbody>
            </table>

            {!previewChanges.length && (
              <div className="empty-state">
                Выберите фамилии, дату и значение.
              </div>
            )}

            {previewChanges.length > 80 && (
              <div className="status-box">Показаны первые 80 изменений.</div>
            )}
          </div>
        </section>
      </section>
      )}
    </main>
  )
}

export default App
