/**
 * Decorative background: a light line-art skyline of civil engineering
 * marvels (suspension bridge, dam, crane, city silhouette) fixed behind the
 * whole app at low opacity. Purely decorative — aria-hidden, no interaction.
 */
export default function Skyline() {
  return (
    <svg
      className="skyline-bg"
      viewBox="0 0 1600 320"
      preserveAspectRatio="xMidYMax slice"
      aria-hidden="true"
      focusable="false"
    >
      {/* horizon */}
      <line x1="0" y1="300" x2="1600" y2="300" stroke="currentColor" strokeWidth="2" />

      {/* dam, far left */}
      <path
        d="M0 300 L0 210 Q90 175 190 210 L190 300 Z"
        fill="currentColor"
        opacity="0.35"
      />
      <path d="M0 210 Q90 175 190 210" fill="none" stroke="currentColor" strokeWidth="2" />
      <path d="M20 240 Q90 215 170 240" fill="none" stroke="currentColor" strokeWidth="1.4" opacity="0.7" />
      <path d="M30 268 Q90 248 160 268" fill="none" stroke="currentColor" strokeWidth="1.4" opacity="0.7" />
      <path d="M190 230 h34 v20 h-34 Z" fill="none" stroke="currentColor" strokeWidth="2" />
      <path d="M230 300 q10 -18 24 -18 t24 18" fill="none" stroke="currentColor" strokeWidth="1.6" opacity="0.6" />

      {/* city silhouette, low buildings between marvels */}
      <g fill="currentColor" opacity="0.22">
        <rect x="270" y="245" width="34" height="55" />
        <rect x="312" y="220" width="26" height="80" />
        <rect x="346" y="260" width="30" height="40" />
        <rect x="1120" y="235" width="30" height="65" />
        <rect x="1156" y="255" width="24" height="45" />
        <rect x="1440" y="225" width="28" height="75" />
        <rect x="1472" y="250" width="34" height="50" />
        <rect x="1510" y="265" width="24" height="35" />
      </g>

      {/* suspension bridge, centre */}
      <g stroke="currentColor" fill="none" strokeWidth="2">
        <line x1="640" y1="120" x2="640" y2="300" />
        <line x1="960" y1="120" x2="960" y2="300" />
        <line x1="622" y1="150" x2="658" y2="150" />
        <line x1="942" y1="150" x2="978" y2="150" />
        <line x1="470" y1="245" x2="1130" y2="245" strokeWidth="3" />
        <path d="M470 245 Q640 130 800 205 Q960 130 1130 245" strokeWidth="2" />
        <path d="M470 245 Q640 300 800 275 Q960 300 1130 245" strokeWidth="1.4" opacity="0.55" />
        {Array.from({ length: 13 }, (_, i) => {
          const x = 490 + i * 50
          return <line key={x} x1={x} y1="245" x2={x} y2="205" strokeWidth="1.2" opacity="0.6" />
        })}
      </g>

      {/* construction crane, right */}
      <g stroke="currentColor" fill="none" strokeWidth="2">
        <line x1="1260" y1="90" x2="1260" y2="300" />
        <line x1="1260" y1="100" x2="1380" y2="100" />
        <line x1="1260" y1="100" x2="1200" y2="120" />
        <line x1="1230" y1="300" x2="1290" y2="300" />
        <line x1="1230" y1="300" x2="1260" y2="240" />
        <line x1="1290" y1="300" x2="1260" y2="240" />
        <line x1="1340" y1="100" x2="1340" y2="150" strokeWidth="1.4" opacity="0.7" />
      </g>
    </svg>
  )
}
