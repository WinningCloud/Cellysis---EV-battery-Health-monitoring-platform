import { useCallback, useEffect, useMemo, useState } from 'react'
import * as XLSX from 'xlsx'
import {
  CartesianGrid,
  Line,
  LineChart,
  ReferenceLine,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from 'recharts'
import './App.css'

const FILE_1_ERROR =
  'File 1 format not recognized. Expected Arbin cycler export format.'
const FILE_2_ERROR =
  'File 2 format not recognized. Expected columns: Duration (sec), mV, mA'
const FILE_2_Q_ERROR =
  'Could not extract discharge capacity from File 2. Ensure the file contains discharge rows (negative mA values).'
const ONE_CYCLE_WARNING =
  'Only 1 unique Cycle_Index detected in File 1. The file was still fully parsed; thousands of rows can belong to one cycle.'

function formatNumber(value, digits = 3) {
  return Number.isFinite(value) ? value.toFixed(digits) : '--'
}

function formatSohPercent(value) {
  if (!Number.isFinite(value)) {
    return '--'
  }
  return `${Math.min(value, 100).toFixed(2)}%`
}

function getSohBand(value) {
  if (!Number.isFinite(value)) {
    return 'unknown'
  }
  if (value >= 80) {
    return 'good'
  }
  if (value >= 60) {
    return 'warn'
  }
  return 'bad'
}

function formatSeconds(totalSeconds) {
  if (!Number.isFinite(totalSeconds) || totalSeconds < 0) {
    return '--'
  }
  const seconds = Math.floor(totalSeconds)
  const hours = String(Math.floor(seconds / 3600)).padStart(2, '0')
  const minutes = String(Math.floor((seconds % 3600) / 60)).padStart(2, '0')
  const remaining = String(seconds % 60).padStart(2, '0')
  return `${hours}:${minutes}:${remaining}`
}

function parseFlexibleNumber(value) {
  if (typeof value === 'number') {
    return Number.isFinite(value) ? value : NaN
  }

  if (typeof value === 'string') {
    const cleaned = value.trim().replace(/,/g, '')
    if (!cleaned) {
      return NaN
    }
    const parsed = Number(cleaned)
    return Number.isFinite(parsed) ? parsed : NaN
  }

  return NaN
}

function parseExcelDate(value) {
  if (value === null || value === undefined) {
    return null
  }

  if (typeof value === 'number' && Number.isFinite(value)) {
    const parts = XLSX.SSF.parse_date_code(value)
    if (!parts) {
      return null
    }
    return new Date(
      parts.y,
      parts.m - 1,
      parts.d,
      parts.H,
      parts.M,
      Math.floor(parts.S),
    )
  }

  const text = String(value).trim()
  if (!text) {
    return null
  }

  const asDate = new Date(text)
  if (!Number.isNaN(asDate.getTime())) {
    return asDate
  }

  return null
}

function parseExcelRows(file) {
  return file.arrayBuffer().then((buffer) => {
    const workbook = XLSX.read(buffer, { type: 'array' })
    const firstSheetName = workbook.SheetNames[0]
    const firstSheet = workbook.Sheets[firstSheetName]
    if (!firstSheet) {
      return []
    }
    return XLSX.utils.sheet_to_json(firstSheet, { defval: null })
  })
}

function normalizeHeader(value) {
  return String(value ?? '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '')
}

function getHeaderIndex(rows) {
  const firstRow = rows[0] ?? {}
  const headerIndex = new Map()

  Object.keys(firstRow).forEach((key) => {
    const normalized = normalizeHeader(key)
    if (!headerIndex.has(normalized)) {
      headerIndex.set(normalized, key)
    }
  })

  return headerIndex
}

function resolveHeaderKey(headerIndex, aliases) {
  for (const alias of aliases) {
    if (headerIndex.has(alias)) {
      return headerIndex.get(alias)
    }
  }
  return null
}

function resolveHeaderByPattern(rows, pattern) {
  const firstRow = rows[0] ?? {}
  const originalHeaders = Object.keys(firstRow)
  const match = originalHeaders.find((header) => pattern.test(normalizeHeader(header)))
  return match ?? null
}

function getWorkbookKind(rows) {
  const headerIndex = getHeaderIndex(rows)
  const hasArbinCycleColumns =
    Boolean(resolveHeaderKey(headerIndex, ['cycleindex'])) &&
    Boolean(resolveHeaderKey(headerIndex, ['dischargecapacityah']))
  const hasSimpleColumns =
    Boolean(resolveHeaderKey(headerIndex, ['durationsec', 'durations'])) &&
    Boolean(resolveHeaderKey(headerIndex, ['mv'])) &&
    Boolean(resolveHeaderKey(headerIndex, ['ma']))

  if (hasArbinCycleColumns) {
    return 'arbin'
  }
  if (hasSimpleColumns) {
    return 'simple'
  }
  return 'unknown'
}

function parseBaselineCapacity(rows) {
  const headerIndex = getHeaderIndex(rows)

  const mvHeader = resolveHeaderKey(headerIndex, ['mv'])
  const maHeader = resolveHeaderKey(headerIndex, ['ma'])

  if (!mvHeader || !maHeader) {
    throw new Error(FILE_2_ERROR)
  }

  const dischargeCapacityHeader =
    resolveHeaderKey(headerIndex, ['dischargecapacityah']) ??
    resolveHeaderByPattern(rows, /^dischargecapacityah?$/)

  if (dischargeCapacityHeader) {
    const dischargeCapacities = rows
      .map((row) => Number(row[dischargeCapacityHeader]))
      .filter((value) => Number.isFinite(value))

    const maxCapacity = dischargeCapacities.length ? Math.max(...dischargeCapacities) : 0

    if (maxCapacity <= 0) {
      throw new Error(FILE_2_Q_ERROR)
    }

    return maxCapacity
  }

  const durationHeader =
    resolveHeaderKey(headerIndex, ['durationsec', 'durations']) ??
    resolveHeaderByPattern(rows, /^duration/)

  if (!durationHeader) {
    throw new Error(FILE_2_ERROR)
  }

  const dischargeRows = rows
    .map((row) => ({
      duration: parseFlexibleNumber(row[durationHeader]),
      currentMilliamp: parseFlexibleNumber(row[maHeader]),
    }))
    .filter((row) => Number.isFinite(row.duration) && Number.isFinite(row.currentMilliamp) && row.currentMilliamp < 0)

  let qInitial = 0
  for (let i = 1; i < dischargeRows.length; i += 1) {
    const previous = dischargeRows[i - 1]
    const current = dischargeRows[i]
    const deltaT = current.duration - previous.duration

    if (deltaT <= 0) {
      continue
    }

    const deltaAh = Math.abs(current.currentMilliamp) * deltaT / (1000 * 3600)
    qInitial += deltaAh
  }

  if (qInitial <= 0) {
    throw new Error(FILE_2_Q_ERROR)
  }

  return qInitial
}

function parseAgedCapacity(rows) {
  const headerIndex = getHeaderIndex(rows)

  const dischargeHeader = resolveHeaderKey(headerIndex, ['dischargecapacityah'])
  const cycleHeader = resolveHeaderKey(headerIndex, ['cycleindex'])
  const stepIndexHeader = resolveHeaderKey(headerIndex, ['stepindex'])
  const testTimeHeader = resolveHeaderKey(headerIndex, ['testtimes'])
  const dateTimeHeader = resolveHeaderKey(headerIndex, ['datetime'])

  if (!dischargeHeader || !cycleHeader) {
    throw new Error(FILE_1_ERROR)
  }

  const cycleMap = new Map()
  const dischargeSeries = []

  rows.forEach((row) => {
    const cycle = parseFlexibleNumber(row[cycleHeader])
    const dischargeCapacity = parseFlexibleNumber(row[dischargeHeader])

    if (Number.isFinite(dischargeCapacity)) {
      dischargeSeries.push(dischargeCapacity)
    }

    if (!Number.isFinite(cycle) || !Number.isFinite(dischargeCapacity)) {
      return
    }

    const existing = cycleMap.get(cycle)
    if (existing === undefined || dischargeCapacity > existing) {
      cycleMap.set(cycle, dischargeCapacity)
    }
  })

  const cycleData = [...cycleMap.entries()]
    .sort((a, b) => a[0] - b[0])
    .map(([cycle, dischargeCapacity]) => ({ cycle, dischargeCapacity }))

  let normalizedCycleData = cycleData

  if (cycleData.length <= 1) {
    // Fallback for exports where Cycle_Index is flat/missing but cycle boundaries can be inferred.
    const inferredCycleMap = new Map()
    let inferredCycle = 1
    let previousDischarge = null
    let previousStepIndex = null
    let previousTestTime = null
    let previousDateTime = null

    rows.forEach((row) => {
      const dischargeCapacity = parseFlexibleNumber(row[dischargeHeader])
      const stepIndex = stepIndexHeader ? parseFlexibleNumber(row[stepIndexHeader]) : NaN
      const testTime = testTimeHeader ? parseFlexibleNumber(row[testTimeHeader]) : NaN
      const dateTime = dateTimeHeader ? parseExcelDate(row[dateTimeHeader]) : null
      if (!Number.isFinite(dischargeCapacity)) {
        return
      }

      if (
        previousDischarge !== null &&
        dischargeCapacity + 1e-6 < previousDischarge
        || (Number.isFinite(stepIndex) && previousStepIndex !== null && stepIndex < previousStepIndex)
        || (Number.isFinite(testTime) && previousTestTime !== null && testTime < previousTestTime)
        || (dateTime && previousDateTime && dateTime.getTime() < previousDateTime.getTime())
      ) {
        inferredCycle += 1
      }

      const existing = inferredCycleMap.get(inferredCycle)
      if (existing === undefined || dischargeCapacity > existing) {
        inferredCycleMap.set(inferredCycle, dischargeCapacity)
      }

      previousDischarge = dischargeCapacity
      if (Number.isFinite(stepIndex)) {
        previousStepIndex = stepIndex
      }
      if (Number.isFinite(testTime)) {
        previousTestTime = testTime
      }
      if (dateTime) {
        previousDateTime = dateTime
      }
    })

    if (inferredCycleMap.size > cycleData.length) {
      normalizedCycleData = [...inferredCycleMap.entries()].map(
        ([cycle, dischargeCapacity]) => ({ cycle, dischargeCapacity }),
      )
    }
  }

  if (!normalizedCycleData.length) {
    throw new Error(FILE_1_ERROR)
  }

  const qCurrent = Math.max(
    ...normalizedCycleData.map((item) => item.dischargeCapacity),
  )

  const resistanceHeader = resolveHeaderKey(headerIndex, ['internalresistanceohm'])

  const resistanceValues = rows
    .map((row) => parseFlexibleNumber(row[resistanceHeader]))
    .filter((value) => Number.isFinite(value) && value !== 0)

  const avgResistanceOhm = resistanceValues.length
    ? resistanceValues.reduce((sum, value) => sum + value, 0) / resistanceValues.length
    : NaN

  const validTestTimes = rows
    .map((row) => parseFlexibleNumber(row[testTimeHeader]))
    .filter((value) => Number.isFinite(value) && value >= 0)

  const maxTestTime = validTestTimes.length ? Math.max(...validTestTimes) : NaN

  const dateTimes = rows
    .map((row) => row[dateTimeHeader])
    .filter((value) => value !== null && value !== undefined && String(value).trim() !== '')
    .map((value) => parseExcelDate(value))
    .filter((value) => value !== null)
    .filter((value) => !Number.isNaN(value.getTime()))

  let dateRange = '--'
  if (dateTimes.length) {
    const minDate = new Date(Math.min(...dateTimes.map((date) => date.getTime())))
    const maxDate = new Date(Math.max(...dateTimes.map((date) => date.getTime())))
    dateRange = `${minDate.toLocaleString()} - ${maxDate.toLocaleString()}`
  }

  return {
    qCurrent,
    cycleData: normalizedCycleData,
    avgResistanceOhm,
    totalDataPoints: rows.length,
    totalCycles: normalizedCycleData.length,
    maxTestTime,
    dateRange,
  }
}

function buildBaselineDebug(rows) {
  const headerIndex = getHeaderIndex(rows)
  const workbookKind = getWorkbookKind(rows)
  const dischargeCapacityHeader =
    resolveHeaderKey(headerIndex, ['dischargecapacityah']) ??
    resolveHeaderByPattern(rows, /^dischargecapacityah?$/)
  const cycleHeader = resolveHeaderKey(headerIndex, ['cycleindex'])
  const testTimeHeader = resolveHeaderKey(headerIndex, ['testtimes'])
  const dateTimeHeader = resolveHeaderKey(headerIndex, ['datetime'])
  const stepIndexHeader = resolveHeaderKey(headerIndex, ['stepindex'])
  const durationHeader =
    resolveHeaderKey(headerIndex, ['durationsec', 'durations']) ??
    resolveHeaderByPattern(rows, /^duration/)
  const maHeader = resolveHeaderKey(headerIndex, ['ma'])

  const directCapacities = dischargeCapacityHeader
    ? rows
        .map((row) => parseFlexibleNumber(row[dischargeCapacityHeader]))
        .filter((value) => Number.isFinite(value))
    : []

  if (directCapacities.length) {
    return {
      workbookKind,
      rowCount: rows.length,
      method: 'Direct max Discharge_Capacity(Ah)',
      dischargeRowCount: directCapacities.length,
      sampleRows: rows.slice(0, 5).map((row, index) => ({
        index: index + 1,
        cycleIndex: cycleHeader ? row[cycleHeader] : '--',
        testTime: testTimeHeader ? row[testTimeHeader] : '--',
        dateTime: dateTimeHeader ? row[dateTimeHeader] : '--',
        stepIndex: stepIndexHeader ? row[stepIndexHeader] : '--',
        dischargeCapacityAh: row[dischargeCapacityHeader],
      })),
    }
  }

  const dischargeRows = rows
    .map((row, index) => ({
      index: index + 1,
      duration: durationHeader ? parseFlexibleNumber(row[durationHeader]) : NaN,
      currentMilliamp: maHeader ? parseFlexibleNumber(row[maHeader]) : NaN,
    }))
    .filter((row) => Number.isFinite(row.duration) && Number.isFinite(row.currentMilliamp) && row.currentMilliamp < 0)

  const cumulativeSamples = []
  let cumulativeAh = 0
  for (let i = 1; i < dischargeRows.length; i += 1) {
    const previous = dischargeRows[i - 1]
    const current = dischargeRows[i]
    const deltaT = current.duration - previous.duration
    if (deltaT <= 0) {
      continue
    }

    const deltaAh = Math.abs(current.currentMilliamp) * deltaT / (1000 * 3600)
    cumulativeAh += deltaAh

    if (cumulativeSamples.length < 5) {
      cumulativeSamples.push({
        fromRow: previous.index,
        toRow: current.index,
        deltaSeconds: deltaT,
        currentMilliamp: current.currentMilliamp,
        deltaAh,
        cumulativeAh,
      })
    }
  }

  return {
    workbookKind,
    rowCount: rows.length,
    method: 'Coulomb counting from negative mA rows',
    dischargeRowCount: dischargeRows.length,
    sampleRows: dischargeRows.slice(0, 5).map((row) => ({
      index: row.index,
      duration: row.duration,
      currentMilliamp: row.currentMilliamp,
    })),
    cumulativeSamples,
  }
}

function buildCurrentDebug(rows) {
  const headerIndex = getHeaderIndex(rows)
  const workbookKind = getWorkbookKind(rows)
  const dischargeHeader = resolveHeaderKey(headerIndex, ['dischargecapacityah'])
  const cycleHeader = resolveHeaderKey(headerIndex, ['cycleindex'])
  const testTimeHeader = resolveHeaderKey(headerIndex, ['testtimes'])
  const dateTimeHeader = resolveHeaderKey(headerIndex, ['datetime'])
  const durationHeader =
    resolveHeaderKey(headerIndex, ['durationsec', 'durations']) ??
    resolveHeaderByPattern(rows, /^duration/)
  const maHeader = resolveHeaderKey(headerIndex, ['ma'])

  const sampleRows = []
  const dischargeRows = []
  const directCapacities = dischargeHeader
    ? rows
        .map((row) => parseFlexibleNumber(row[dischargeHeader]))
        .filter((value) => Number.isFinite(value))
    : []

  rows.forEach((row, index) => {
    const duration = durationHeader ? parseFlexibleNumber(row[durationHeader]) : NaN
    const currentMilliamp = maHeader ? parseFlexibleNumber(row[maHeader]) : NaN
    const dischargeCapacity = dischargeHeader ? parseFlexibleNumber(row[dischargeHeader]) : NaN

    if (Number.isFinite(duration) && Number.isFinite(currentMilliamp) && currentMilliamp < 0) {
      dischargeRows.push({
        index: index + 1,
        duration,
        currentMilliamp,
      })
    }

    if (sampleRows.length < 5) {
      sampleRows.push({
        row: index + 1,
        duration: Number.isFinite(duration) ? duration : '--',
        currentMilliamp: Number.isFinite(currentMilliamp) ? currentMilliamp : '--',
        cycleIndex: cycleHeader ? row[cycleHeader] : '--',
        testTime: testTimeHeader ? row[testTimeHeader] : '--',
        dateTime: dateTimeHeader ? row[dateTimeHeader] : '--',
        dischargeCapacityAh: Number.isFinite(dischargeCapacity) ? dischargeCapacity : '--',
      })
    }
  })

  let qCurrent = NaN
  let cumulativeAh = 0
  const cumulativeSamples = []

  if (directCapacities.length) {
    qCurrent = Math.max(...directCapacities)
  } else {
    for (let i = 1; i < dischargeRows.length; i += 1) {
      const previous = dischargeRows[i - 1]
      const current = dischargeRows[i]
      const deltaT = current.duration - previous.duration

      if (deltaT <= 0) {
        continue
      }

      const deltaAh = Math.abs(current.currentMilliamp) * deltaT / (1000 * 3600)
      cumulativeAh += deltaAh

      if (cumulativeSamples.length < 5) {
        cumulativeSamples.push({
          fromRow: previous.index,
          toRow: current.index,
          deltaSeconds: deltaT,
          currentMilliamp: current.currentMilliamp,
          deltaAh,
          cumulativeAh,
        })
      }
    }

    qCurrent = cumulativeAh
  }

  return {
    workbookKind,
    rowCount: rows.length,
    method: directCapacities.length
      ? 'Direct max Discharge_Capacity(Ah)'
      : 'Coulomb counting from negative mA rows',
    qCurrent,
    dischargeRowCount: directCapacities.length || dischargeRows.length,
    sampleRows,
    cumulativeSamples,
  }
}

function App() {
  const [isDarkMode, setIsDarkMode] = useState(false)
  const [baselineFile, setBaselineFile] = useState(null)
  const [testFile, setTestFile] = useState(null)
  const [analysis, setAnalysis] = useState(null)
  const [isAnalyzing, setIsAnalyzing] = useState(false)
  const [errorMessage, setErrorMessage] = useState('')
  const [warningMessage, setWarningMessage] = useState('')
  const [showRawMetrics, setShowRawMetrics] = useState(false)
  const [showDebug, setShowDebug] = useState(false)
  const [dragTarget, setDragTarget] = useState(null)

  useEffect(() => {
    document.documentElement.classList.toggle('dark', isDarkMode)
  }, [isDarkMode])

  const handleFilePick = useCallback((file, type) => {
    if (!file) {
      return
    }

    if (!file.name.toLowerCase().endsWith('.xlsx')) {
      setErrorMessage('Please upload .xlsx files only.')
      return
    }

    setErrorMessage('')
    setWarningMessage('')
    setAnalysis(null)

    if (type === 'baseline') {
      setBaselineFile(file)
    } else {
      setTestFile(file)
    }
  }, [])

  const analyzeFiles = useCallback(async () => {
    if (!baselineFile || !testFile) {
      return
    }

    try {
      setIsAnalyzing(true)
      setErrorMessage('')
      setWarningMessage('')

      const [baselineRows, testRows] = await Promise.all([
        parseExcelRows(baselineFile),
        parseExcelRows(testFile),
      ])

      const baselineKind = getWorkbookKind(baselineRows)
      const testKind = getWorkbookKind(testRows)

      const baselineDebug = buildBaselineDebug(baselineRows)
      const currentDebug = buildCurrentDebug(testRows)

      const arbinRows = baselineKind === 'arbin' ? baselineRows : testRows
      const simpleRows = baselineKind === 'simple' ? baselineRows : testRows

      if (baselineKind === 'unknown' || testKind === 'unknown') {
        throw new Error(
          'Unable to detect one or both workbook types. One file must be an Arbin export and the other must be a Duration/mV/mA log.',
        )
      }

      const initialMetrics = parseAgedCapacity(arbinRows)
      const qInitial = initialMetrics.qCurrent
      const qCurrent = parseBaselineCapacity(simpleRows)
      const cycleData = initialMetrics.cycleData

      const sohRaw = (qCurrent / qInitial) * 100
      const sohCapped = Math.min(sohRaw, 100)
      const capacityFade = qInitial - qCurrent

      const cycleTable = cycleData.map((item) => {
        const cycleSohRaw = (item.dischargeCapacity / qInitial) * 100
        return {
          cycle: item.cycle,
          dischargeCapacity: item.dischargeCapacity,
          sohRaw: cycleSohRaw,
          sohDisplay: Math.min(cycleSohRaw, 100),
          internalResistanceMilliOhm: Number.isFinite(initialMetrics.avgResistanceOhm)
            ? initialMetrics.avgResistanceOhm * 1000
            : NaN,
        }
      })

      if (cycleData.length === 1) {
        setWarningMessage(ONE_CYCLE_WARNING)
      }

      setAnalysis({
        qInitial,
        qCurrent,
        sohRaw,
        sohCapped,
        capacityFade,
        avgResistanceOhm: initialMetrics.avgResistanceOhm,
        cycleData,
        cycleTable,
        rawMetrics: {
          totalDataPoints: initialMetrics.totalDataPoints,
          totalCycles: initialMetrics.totalCycles,
          maxTestTime: initialMetrics.maxTestTime,
          dateRange: initialMetrics.dateRange,
        },
        debug: {
          baseline: baselineDebug,
          current: currentDebug,
          detected: {
            baselineKind,
            testKind,
            arbinRowsCount: arbinRows.length,
            simpleRowsCount: simpleRows.length,
          },
        },
      })
    } catch (error) {
      setAnalysis(null)
      setErrorMessage(error instanceof Error ? error.message : 'Failed to analyze files.')
    } finally {
      setIsAnalyzing(false)
    }
  }, [baselineFile, testFile])

  const sohBand = useMemo(() => {
    if (!analysis) {
      return 'unknown'
    }
    return getSohBand(analysis.sohCapped)
  }, [analysis])

  const tooltipStyles = useMemo(() => {
    return {
      backgroundColor: isDarkMode ? '#0f172a' : '#f8fafc',
      border: isDarkMode ? '1px solid #1e293b' : '1px solid #cbd5e1',
      borderRadius: '12px',
      boxShadow: isDarkMode
        ? '0 12px 28px rgba(2, 6, 23, 0.5)'
        : '0 12px 28px rgba(15, 23, 42, 0.15)',
      color: isDarkMode ? '#e2e8f0' : '#0f172a',
    }
  }, [isDarkMode])

  const canAnalyze = Boolean(baselineFile && testFile)

  const cardStateClass = useCallback(
    (type) => {
      const isLoaded = type === 'baseline' ? baselineFile : testFile
      const isDragging = dragTarget === type
      return [
        'upload-card',
        isLoaded ? 'upload-card--loaded' : '',
        isDragging ? 'upload-card--drag' : '',
      ]
        .filter(Boolean)
        .join(' ')
    },
    [baselineFile, testFile, dragTarget],
  )

  const handleDrop = useCallback((event, type) => {
    event.preventDefault()
    setDragTarget(null)
    const file = event.dataTransfer?.files?.[0]
    handleFilePick(file, type)
  }, [handleFilePick])

  return (
    <div className="app-shell">
      <div className="app-backdrop" aria-hidden="true" />

      <main className="app-container">
        <header className="top-header card-surface">
          <div className="top-header__identity">
            <h1 className="top-header__title">Cellysis</h1>
            <p className="top-header__subtitle">Battery State of Health Analyzer</p>
          </div>
          <button
            type="button"
            onClick={() => setIsDarkMode((prev) => !prev)}
            className="theme-toggle"
            aria-label="Toggle dark mode"
          >
            <span className="theme-toggle__knob" aria-hidden="true">{isDarkMode ? '☾' : '☀'}</span>
            <span>{isDarkMode ? 'Dark' : 'Light'} Mode</span>
          </button>
          <div className="top-header__accent" aria-hidden="true" />
        </header>

        <section className="card-surface upload-section">
          <div className="section-title-wrap">
            <h2 className="section-title">Step 1: Upload Test Files</h2>
            <p className="section-subtitle">Provide baseline and current workbook exports for SOH analysis.</p>
          </div>

          <div className="upload-grid">
            <label
              className={cardStateClass('baseline')}
              onDragOver={(event) => {
                event.preventDefault()
                setDragTarget('baseline')
              }}
              onDragLeave={() => setDragTarget(null)}
              onDrop={(event) => handleDrop(event, 'baseline')}
            >
              <input
                type="file"
                accept=".xlsx"
                className="hidden"
                onChange={(event) => handleFilePick(event.target.files?.[0], 'baseline')}
              />
              <div className="upload-card__head">
                <span className="upload-icon" aria-hidden="true">⇪</span>
                <div>
                  <p className="upload-card__title">Initial Capacity File</p>
                  <p className="upload-card__subtitle">Arbin cycler export workbook</p>
                </div>
              </div>

              <div className="upload-card__picker">
                <span>Drop .xlsx here or click to browse</span>
              </div>

              <div className="upload-status">
                {baselineFile ? (
                  <>
                    <span className="status-check" aria-hidden="true">✓</span>
                    <span className="upload-filename">{baselineFile.name}</span>
                  </>
                ) : (
                  <span className="upload-empty">No file selected</span>
                )}
              </div>
            </label>

            <label
              className={cardStateClass('test')}
              onDragOver={(event) => {
                event.preventDefault()
                setDragTarget('test')
              }}
              onDragLeave={() => setDragTarget(null)}
              onDrop={(event) => handleDrop(event, 'test')}
            >
              <input
                type="file"
                accept=".xlsx"
                className="hidden"
                onChange={(event) => handleFilePick(event.target.files?.[0], 'test')}
              />
              <div className="upload-card__head">
                <span className="upload-icon" aria-hidden="true">⇪</span>
                <div>
                  <p className="upload-card__title">Current Capacity File</p>
                  <p className="upload-card__subtitle">Duration/mV/mA workbook</p>
                </div>
              </div>

              <div className="upload-card__picker">
                <span>Drop .xlsx here or click to browse</span>
              </div>

              <div className="upload-status">
                {testFile ? (
                  <>
                    <span className="status-check" aria-hidden="true">✓</span>
                    <span className="upload-filename">{testFile.name}</span>
                  </>
                ) : (
                  <span className="upload-empty">No file selected</span>
                )}
              </div>
            </label>
          </div>

          {canAnalyze && (
            <div className="analyze-wrap">
              <button
                type="button"
                onClick={analyzeFiles}
                disabled={isAnalyzing}
                className="analyze-button"
              >
                {isAnalyzing ? 'Analyzing...' : 'Analyze'}
              </button>
            </div>
          )}

          {errorMessage && <p className="status-alert status-alert--error">{errorMessage}</p>}
          {warningMessage && !errorMessage && (
            <p className="status-alert status-alert--warn">{warningMessage}</p>
          )}
        </section>

        {analysis && (
          <section className="results-flow">
            <div className="kpi-grid">
              <article className="card-surface kpi-card">
                <p className="kpi-label">SOH</p>
                <p className={`kpi-value kpi-value--${sohBand}`}>{formatSohPercent(analysis.sohRaw)}</p>
                <p className="kpi-caption">State of health ratio</p>
              </article>

              <article className="card-surface kpi-card">
                <p className="kpi-label">Initial Capacity (Q_initial)</p>
                <p className="kpi-value">{formatNumber(analysis.qInitial)} Ah</p>
                <p className="kpi-caption">Baseline</p>
              </article>

              <article className="card-surface kpi-card">
                <p className="kpi-label">Current Capacity (Q_current)</p>
                <p className="kpi-value">{formatNumber(analysis.qCurrent)} Ah</p>
                <p className="kpi-caption">Measured</p>
              </article>

              <article className="card-surface kpi-card">
                <p className="kpi-label">Avg Internal Resistance</p>
                <p className="kpi-value">
                  {Number.isFinite(analysis.avgResistanceOhm)
                    ? `${formatNumber(analysis.avgResistanceOhm * 1000, 2)} mΩ`
                    : '--'}
                </p>
                <p className="kpi-caption">Cell average resistance</p>
              </article>
            </div>

            <article className="card-surface chart-card">
              <div className="chart-head">
                <h3 className="section-title">Discharge Capacity per Cycle</h3>
                <p className="chart-caption">Capacity Fade: {formatNumber(analysis.capacityFade)} Ah</p>
              </div>

              <div className="chart-wrap">
                <ResponsiveContainer width="100%" height="100%">
                  <LineChart data={analysis.cycleData} margin={{ top: 20, right: 25, left: 10, bottom: 20 }}>
                    <CartesianGrid strokeDasharray="4 4" stroke={isDarkMode ? '#334155' : '#cbd5e1'} />
                    <XAxis
                      dataKey="cycle"
                      tick={{ fill: isDarkMode ? '#cbd5e1' : '#334155', fontSize: 12 }}
                      tickLine={false}
                      axisLine={{ stroke: isDarkMode ? '#334155' : '#cbd5e1' }}
                      label={{
                        value: 'Cycle Number',
                        position: 'insideBottom',
                        offset: -10,
                        fill: isDarkMode ? '#94a3b8' : '#475569',
                      }}
                    />
                    <YAxis
                      tick={{ fill: isDarkMode ? '#cbd5e1' : '#334155', fontSize: 12 }}
                      tickLine={false}
                      axisLine={{ stroke: isDarkMode ? '#334155' : '#cbd5e1' }}
                      label={{
                        value: 'Discharge Capacity (Ah)',
                        angle: -90,
                        position: 'insideLeft',
                        fill: isDarkMode ? '#94a3b8' : '#475569',
                      }}
                    />
                    <Tooltip
                      contentStyle={tooltipStyles}
                      formatter={(value) => `${Number(value).toFixed(3)} Ah`}
                      labelFormatter={(label) => `Cycle ${label}`}
                    />
                    <ReferenceLine
                      y={analysis.qInitial}
                      stroke="#dc2626"
                      strokeDasharray="8 5"
                      label={{ value: 'Baseline (100% SOH)', fill: '#dc2626', position: 'right' }}
                    />
                    <ReferenceLine
                      y={analysis.qInitial * 0.8}
                      stroke="#f59e0b"
                      strokeDasharray="8 5"
                      label={{ value: 'EOL threshold (80% SOH)', fill: '#f59e0b', position: 'right' }}
                    />
                    <Line
                      type="monotone"
                      dataKey="dischargeCapacity"
                      stroke="#0d9488"
                      strokeWidth={2.6}
                      dot={{ r: 1.8, strokeWidth: 0, fill: '#0d9488' }}
                      activeDot={{ r: 4.5, fill: '#14b8a6', stroke: '#0f172a', strokeWidth: 1 }}
                    />
                  </LineChart>
                </ResponsiveContainer>
              </div>
            </article>

            <article className="card-surface table-card">
              <h3 className="section-title">Cycle Table</h3>
              <div className="table-scroll">
                <table className="cycle-table">
                  <thead>
                    <tr>
                      <th>Cycle</th>
                      <th>Discharge Capacity (Ah)</th>
                      <th>SOH (%)</th>
                      <th>Internal Resistance (mΩ)</th>
                    </tr>
                  </thead>
                  <tbody>
                    {analysis.cycleTable.map((row) => {
                      const band = getSohBand(row.sohDisplay)
                      return (
                        <tr key={row.cycle} className={`row-band row-band--${band}`}>
                          <td>{row.cycle}</td>
                          <td>{formatNumber(row.dischargeCapacity)}</td>
                          <td>{formatSohPercent(row.sohRaw)}</td>
                          <td>
                            {Number.isFinite(row.internalResistanceMilliOhm)
                              ? formatNumber(row.internalResistanceMilliOhm, 2)
                              : '--'}
                          </td>
                        </tr>
                      )
                    })}
                  </tbody>
                </table>
              </div>
            </article>

            <article className="card-surface collapse-panel">
              <button
                type="button"
                onClick={() => setShowRawMetrics((prev) => !prev)}
                className="collapse-toggle"
              >
                <span>{showRawMetrics ? '▼' : '▶'}</span>
                <span>Raw Metrics</span>
              </button>

              {showRawMetrics && (
                <div className="collapse-content metrics-grid">
                  <div className="mini-metric">
                    <p>Total data points</p>
                    <strong>{analysis.rawMetrics.totalDataPoints}</strong>
                  </div>
                  <div className="mini-metric">
                    <p>Unique Cycle_Index values</p>
                    <strong>{analysis.rawMetrics.totalCycles}</strong>
                  </div>
                  <div className="mini-metric">
                    <p>Test duration</p>
                    <strong>{formatSeconds(analysis.rawMetrics.maxTestTime)}</strong>
                  </div>
                  <div className="mini-metric">
                    <p>Date range</p>
                    <strong>{analysis.rawMetrics.dateRange}</strong>
                  </div>
                </div>
              )}
            </article>

            <article className="card-surface collapse-panel">
              <button
                type="button"
                onClick={() => setShowDebug((prev) => !prev)}
                className="collapse-toggle"
              >
                <span>{showDebug ? '▼' : '▶'}</span>
                <span>Extraction Verification</span>
              </button>

              {showDebug && (
                <div className="collapse-content debug-layout">
                  <div className="debug-card-grid">
                    <div className="mini-metric mono-metric">
                      <p>Detected workbook types</p>
                      <strong>
                        baseline={analysis.debug.detected.baselineKind} | current={analysis.debug.detected.testKind}
                      </strong>
                    </div>
                    <div className="mini-metric mono-metric">
                      <p>Rows read</p>
                      <strong>
                        file1={analysis.debug.baseline.rowCount} | file2={analysis.debug.current.rowCount}
                      </strong>
                    </div>
                    <div className="mini-metric mono-metric">
                      <p>Q_initial source and value</p>
                      <strong>{analysis.debug.baseline.method} | {formatNumber(analysis.qInitial)} Ah</strong>
                    </div>
                    <div className="mini-metric mono-metric">
                      <p>Q_current source and value</p>
                      <strong>
                        {analysis.debug.current.method} | {formatNumber(analysis.debug.current.qCurrent)} Ah
                      </strong>
                    </div>
                    <div className="mini-metric mono-metric">
                      <p>Rows used for extraction</p>
                      <strong>
                        baseline={analysis.debug.baseline.dischargeRowCount} | current={analysis.debug.current.dischargeRowCount}
                      </strong>
                    </div>
                  </div>

                  <div className="debug-rows-grid">
                    <div className="debug-sample-panel">
                      <h4>Sample Parsed Rows - File 1</h4>
                      <div className="debug-sample-list">
                        {analysis.debug.baseline.sampleRows.map((row) => (
                          <div key={row.index} className="sample-chip">
                            <div>row: {row.index}</div>
                            {analysis.debug.baseline.workbookKind === 'arbin' ? (
                              <>
                                <div>cycle: {String(row.cycleIndex ?? '--')}</div>
                                <div>test_time: {String(row.testTime ?? '--')}</div>
                                <div>step_index: {String(row.stepIndex ?? '--')}</div>
                              </>
                            ) : (
                              <>
                                <div>duration: {String(row.duration ?? '--')}</div>
                                <div>mA: {String(row.currentMilliamp ?? '--')}</div>
                              </>
                            )}
                            <div>q: {String(row.dischargeCapacityAh ?? '--')}</div>
                          </div>
                        ))}
                      </div>
                    </div>

                    <div className="debug-sample-panel">
                      <h4>Sample Parsed Rows - File 2</h4>
                      <div className="debug-sample-list">
                        {analysis.debug.current.sampleRows.map((row) => (
                          <div key={row.row} className="sample-chip">
                            <div>row: {row.row}</div>
                            {analysis.debug.current.workbookKind === 'simple' ? (
                              <>
                                <div>duration: {String(row.duration ?? '--')}</div>
                                <div>mA: {String(row.currentMilliamp ?? '--')}</div>
                              </>
                            ) : (
                              <>
                                <div>cycle: {String(row.cycleIndex ?? '--')}</div>
                                <div>test_time: {String(row.testTime ?? '--')}</div>
                                <div>date_time: {String(row.dateTime ?? '--')}</div>
                              </>
                            )}
                            <div>q: {String(row.dischargeCapacityAh ?? '--')}</div>
                          </div>
                        ))}
                      </div>
                    </div>
                  </div>
                </div>
              )}
            </article>
          </section>
        )}
      </main>
    </div>
  )
}

export default App
