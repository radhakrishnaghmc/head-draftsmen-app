import React from 'react'
import { createRoot } from 'react-dom/client'
import AuthGate from './AuthGate'
import './styles.css'

const container = document.getElementById('root')!
createRoot(container).render(
  <React.StrictMode>
    <AuthGate />
  </React.StrictMode>
)
