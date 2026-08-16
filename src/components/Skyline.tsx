import charminar from '../assets/charminar-line-art.png'
import thousandPillarTemple from '../assets/thousand-pillar-temple-line-art.png'

/**
 * Decorative background: very light line-art of Telangana landmarks (Charminar,
 * the Thousand Pillar Temple) fixed behind the workspace. The line art is
 * derived from real reference photographs (edge-detected, not freehand-drawn),
 * shown at very low opacity. Purely decorative — aria-hidden, no interaction.
 */
export default function Skyline() {
  return (
    <div className="skyline-bg" aria-hidden="true">
      <img src={charminar} alt="" className="skyline-charminar" />
      <img src={thousandPillarTemple} alt="" className="skyline-temple" />
    </div>
  )
}
