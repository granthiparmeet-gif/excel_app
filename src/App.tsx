import { useCallback, useEffect, useMemo, useState, type ClipboardEvent, type KeyboardEvent } from 'react'

const columnOrder = [
  'Keyword',
  'Prefix',
  'Suffix',
  'Middle',
  'City',
  'FirstName',
  '3 letter',
  '4 letter',
  'Extensions',
] as const

type ColumnKey = (typeof columnOrder)[number]
type RowData = Record<ColumnKey, string>

// Build an empty row so every addition uses the same shape and stays easy to extend later.
const createEmptyRow = (): RowData =>
  columnOrder.reduce((record, column) => {
    record[column] = ''
    return record
  }, {} as RowData)

const INITIAL_ROW_COUNT = 8
const STORAGE_KEY = 'excel_worksheet_data'

const normalizeValue = (value: string) => value.trim().toLowerCase()

function App() {
  const [rows, setRows] = useState<RowData[]>(() => {
    try {
      const raw = localStorage.getItem(STORAGE_KEY)
      if (!raw) {
        return Array.from({ length: INITIAL_ROW_COUNT }, () => createEmptyRow())
      }
      const parsed = JSON.parse(raw) as Record<string, string>[]

      if (!Array.isArray(parsed) || !parsed.length) {
        return Array.from({ length: INITIAL_ROW_COUNT }, () => createEmptyRow())
      }

      return parsed.map((entry) => {
        const row: RowData = createEmptyRow()
        columnOrder.forEach((column) => {
          row[column] = String(entry[column] ?? '').trim()
        })
        return row
      })
    } catch (error) {
      console.warn('Failed to load saved worksheet data, falling back to empty rows.', error)
      return Array.from({ length: INITIAL_ROW_COUNT }, () => createEmptyRow())
    }
  })
  const [editingCell, setEditingCell] = useState<{ row: number; colIdx: number } | null>(null)

  useEffect(() => {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(rows))
    } catch (error) {
      console.warn('Failed to persist worksheet data', error)
    }
  }, [rows])

  // Maintain a quick lookup of normalized cell values so duplicate highlights stay fast.
  const duplicateLookup = useMemo(() => {
    const lookup = new Map<string, Array<{ row: number; colIdx: number }>>()

    rows.forEach((row, rowIdx) => {
      columnOrder.forEach((column, colIdx) => {
        const normalized = normalizeValue(row[column])
        if (!normalized) return
        if (!lookup.has(normalized)) {
          lookup.set(normalized, [])
        }
        lookup.get(normalized)!.push({ row: rowIdx, colIdx })
      })
    })

    return lookup
  }, [rows])

  // Automatically grow the dataset when navigation reaches past the last row.
  const ensureRowExists = useCallback((targetRow: number) => {
    setRows((prev) => {
      if (targetRow < prev.length) {
        return prev
      }
      const rowsToAdd = targetRow - prev.length + 1
      return [...prev, ...Array.from({ length: rowsToAdd }, () => createEmptyRow())]
    })
  }, [])

  const focusCell = useCallback(
    (targetRow: number, targetCol: number) => {
      if (targetRow < 0) {
        return
      }
      ensureRowExists(targetRow)
      const boundedCol = Math.min(Math.max(targetCol, 0), columnOrder.length - 1)
      setEditingCell({ row: targetRow, colIdx: boundedCol })
    },
    [ensureRowExists],
  )

  const addRow = useCallback(() => {
    setRows((prev) => [...prev, createEmptyRow()])
  }, [])

  const updateCellValue = useCallback((rowIdx: number, column: ColumnKey, value: string) => {
    setRows((prev) => {
      const next = [...prev]
      next[rowIdx] = { ...next[rowIdx], [column]: value }
      return next
    })
  }, [])

  const applyPasteData = useCallback(
    (startRow: number, startCol: number, grid: string[][]) => {
      setRows((prev) => {
        const next = [...prev]

        grid.forEach((gridRow, rowOffset) => {
          const targetRow = startRow + rowOffset
          while (targetRow >= next.length) {
            next.push(createEmptyRow())
          }

          const updatedRow = { ...next[targetRow] }

          gridRow.forEach((value, colOffset) => {
            const targetCol = startCol + colOffset
            if (targetCol < 0 || targetCol >= columnOrder.length) {
              return
            }
            updatedRow[columnOrder[targetCol]] = value
          })

          next[targetRow] = updatedRow
        })

        return next
      })
    },
    [],
  )

  // Keyboard navigation mirrors Excel: Enter moves down, Tab moves right (with wrap/Shift support).
  const handleKeyDown = useCallback(
    (event: KeyboardEvent<HTMLInputElement>, rowIdx: number, colIdx: number) => {
      if (event.key === 'Enter') {
        event.preventDefault()
        focusCell(rowIdx + 1, colIdx)
        return
      }

      if (event.key === 'Tab') {
        event.preventDefault()
        const direction = event.shiftKey ? -1 : 1
        let nextCol = colIdx + direction
        let nextRow = rowIdx

        if (nextCol >= columnOrder.length) {
          nextCol = 0
          nextRow += 1
        } else if (nextCol < 0) {
          nextCol = columnOrder.length - 1
          nextRow = Math.max(0, rowIdx - 1)
        }

        focusCell(nextRow, nextCol)
      }
    },
    [focusCell],
  )

  const handlePaste = useCallback(
    (event: ClipboardEvent<HTMLInputElement>, rowIdx: number, colIdx: number) => {
      event.preventDefault()
      const text = event.clipboardData.getData('text/plain')
      if (!text) {
        return
      }

      const rows = text.split(/\r\n|\n|\r/)
      const parsed = rows.map((row) => row.split('\t'))

      applyPasteData(rowIdx, colIdx, parsed)
    },
    [applyPasteData],
  )

  const isDuplicateCell = (rowIdx: number, column: ColumnKey) => {
    const normalized = normalizeValue(rows[rowIdx][column])
    if (!normalized) {
      return false
    }
    const matches = duplicateLookup.get(normalized)
    return Boolean(matches && matches.length > 1)
  }

  // Keep the grid stable by rendering a non-breaking space in empty cells.
  const cellContent = (value: string) => (value ? value : '\u00a0')

  return (
    <div className="bg-slate-50 px-4">
      <div className="mx-auto w-full max-w-6xl space-y-6 rounded-xl bg-white p-6 shadow-xl ring-1 ring-slate-200">
        <div className="flex flex-wrap items-center justify-between gap-3">
          <div>
            <h1 className="text-2xl font-semibold text-slate-900">Excel-style Worksheet</h1>
            <p className="text-sm text-slate-500">
              Inline edits, Tab/Enter navigation, and real-time duplicate highlighting.
            </p>
          </div>
          <button
            type="button"
            onClick={addRow}
            className="rounded border border-slate-300 bg-slate-100 px-4 py-2 text-sm font-medium text-slate-700 transition hover:border-slate-400 hover:bg-slate-200"
          >
            + Add Row
          </button>
        </div>

        <div className="overflow-auto">
          <div
            className="grid min-w-full rounded-sm bg-white text-sm text-slate-800 shadow-[0_1px_4px_rgba(15,23,42,0.08)]"
            style={{
              gridTemplateColumns: `48px repeat(${columnOrder.length}, minmax(130px,1fr))`,
            }}
          >
            <div className="border-b border-r border-slate-200 bg-slate-100 px-2 py-3 text-xs font-semibold uppercase tracking-wider text-slate-500">
              #
            </div>
            {columnOrder.map((column) => (
              <div
                key={column}
                className="flex items-center border-b border-r border-slate-200 bg-slate-100 px-3 py-3 text-xs font-semibold uppercase tracking-widest text-slate-600"
              >
                {column}
              </div>
            ))}

            {rows.map((row, rowIdx) => (
                  <div
                    key={`row-${rowIdx}`}
                    className="contents"
                  >
                    <div className="border-b border-r border-slate-200 bg-slate-50 px-2 py-3 text-sm font-medium text-slate-500">
                      {rowIdx + 1}
                    </div>

                    {columnOrder.map((column, colIdx) => {
                  const isEditing =
                    editingCell?.row === rowIdx && editingCell.colIdx === colIdx
                  const duplicate = isDuplicateCell(rowIdx, column)
                  const focusRing = isEditing
                    ? 'ring-2 ring-inset ring-blue-400/70'
                    : 'focus-visible:ring-2 focus-visible:ring-inset focus-visible:ring-sky-500'

                  return (
                      <div
                        key={`${rowIdx}-${column}`}
                        className={`border-b border-r border-slate-200 ${
                          duplicate ? 'bg-yellow-100' : 'bg-white'
                        }`}
                        onClick={() => focusCell(rowIdx, colIdx)}
                      >
                      {isEditing ? (
                        <input
                          data-cell-input
                          value={row[column]}
                          onChange={(event) => updateCellValue(rowIdx, column, event.target.value)}
                          onBlur={(event) => {
                            const relatedTarget = event.relatedTarget as HTMLElement | null
                            if (!relatedTarget?.hasAttribute('data-cell-input')) {
                              setEditingCell(null)
                            }
                          }}
                          onKeyDown={(event) => handleKeyDown(event, rowIdx, colIdx)}
                          onPaste={(event) => handlePaste(event, rowIdx, colIdx)}
                          autoFocus
                          className={`h-10 w-full border-none bg-transparent px-3 text-left outline-none ${focusRing}`}
                        />
                      ) : (
                        <div
                          className={`h-10 cursor-text px-3 leading-10 ${focusRing}`}
                          onFocus={() => focusCell(rowIdx, colIdx)}
                          role="presentation"
                        >
                          {cellContent(row[column])}
                        </div>
                      )}
                    </div>
                  )
                })}
              </div>
            ))}
          </div>
        </div>
        <p className="text-xs text-slate-500">
          Values that match (case-insensitive, trimmed) earlier entries are softly highlighted, just
          like Excel duplication hints.
        </p>
      </div>
    </div>
  )
}

export default App
