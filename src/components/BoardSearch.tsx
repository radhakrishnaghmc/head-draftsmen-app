import { IconSearch } from './Icons'

/** Search box shared by the reminders and the Forwarding / QC letter boards (see core/boardSearch.ts). */
export default function BoardSearch({ value, onChange, count, total }: { value: string; onChange: (v: string) => void; count: number; total: number }) {
  return (
    <div className="board-search">
      <IconSearch />
      <input
        type="search"
        value={value}
        onChange={(e) => onChange(e.target.value)}
        placeholder="Search by name of the work or agency name"
        aria-label="Search by name of the work or agency name"
      />
      {value.trim() && (
        <span className="board-search-count">
          {count} of {total}
        </span>
      )}
    </div>
  )
}
