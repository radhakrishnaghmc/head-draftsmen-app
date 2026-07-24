import { useRef, useState } from 'react'
import { IconFolder, IconImage, IconTrash } from './Icons'

interface Photo {
  id: string
  name: string
  dataUrl: string
}

function nextId(): string {
  return `photo-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`
}

function readAsDataUrl(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = () => resolve(reader.result as string)
    reader.onerror = () => reject(reader.error)
    reader.readAsDataURL(file)
  })
}

/**
 * Upload site photos and arrange them into the order they should appear in
 * the estimate — drag a tile to reorder; its number badge updates to match.
 */
export default function UploadPhotosTab() {
  const [photos, setPhotos] = useState<Photo[]>([])
  const [dragIndex, setDragIndex] = useState<number | null>(null)
  const [overIndex, setOverIndex] = useState<number | null>(null)
  const fileInputRef = useRef<HTMLInputElement>(null)

  async function handleFiles(fileList: FileList | null) {
    if (!fileList || fileList.length === 0) return
    const files = Array.from(fileList).filter((f) => f.type.startsWith('image/'))
    const added = await Promise.all(
      files.map(async (file) => ({ id: nextId(), name: file.name, dataUrl: await readAsDataUrl(file) }))
    )
    setPhotos((prev) => [...prev, ...added])
  }

  function removePhoto(id: string) {
    setPhotos((prev) => prev.filter((p) => p.id !== id))
  }

  function handleDragStart(e: React.DragEvent, index: number) {
    setDragIndex(index)
    e.dataTransfer.effectAllowed = 'move'
  }

  function handleDragOver(e: React.DragEvent, index: number) {
    e.preventDefault()
    e.dataTransfer.dropEffect = 'move'
    if (index !== overIndex) setOverIndex(index)
  }

  function handleDrop(e: React.DragEvent, index: number) {
    e.preventDefault()
    if (dragIndex !== null && dragIndex !== index) {
      setPhotos((prev) => {
        const next = [...prev]
        const [moved] = next.splice(dragIndex, 1)
        next.splice(index, 0, moved)
        return next
      })
    }
    setDragIndex(null)
    setOverIndex(null)
  }

  function handleDragEnd() {
    setDragIndex(null)
    setOverIndex(null)
  }

  return (
    <div className="card">
      <div className="empty">
        <IconImage />
        <p>
          {photos.length === 0
            ? 'Upload site photos to arrange them for the estimate.'
            : `${photos.length} photo${photos.length === 1 ? '' : 's'} added — drag to reorder.`}
        </p>
        <div className="boq-actions">
          <button className="primary" onClick={() => fileInputRef.current?.click()}>
            <IconFolder /> Upload Photos
          </button>
          <input
            ref={fileInputRef}
            type="file"
            accept="image/*"
            multiple
            style={{ display: 'none' }}
            onChange={(e) => {
              void handleFiles(e.target.files)
              e.target.value = ''
            }}
          />
        </div>
      </div>

      {photos.length > 0 && (
        <div className="photo-tile-grid">
          {photos.map((photo, i) => (
            <div
              key={photo.id}
              className={[
                'photo-tile-card',
                dragIndex === i ? 'dragging' : '',
                overIndex === i && dragIndex !== null && dragIndex !== i ? 'drag-over' : ''
              ]
                .filter(Boolean)
                .join(' ')}
              draggable
              onDragStart={(e) => handleDragStart(e, i)}
              onDragOver={(e) => handleDragOver(e, i)}
              onDrop={(e) => handleDrop(e, i)}
              onDragEnd={handleDragEnd}
            >
              <span className="photo-tile-number">{i + 1}</span>
              <button className="photo-tile-remove" title="Remove" onClick={() => removePhoto(photo.id)}>
                <IconTrash />
              </button>
              <img className="photo-tile-img" src={photo.dataUrl} alt={photo.name} />
              <div className="photo-tile-name" title={photo.name}>
                {photo.name}
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  )
}
