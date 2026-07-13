import React from "react"
import { createRoot } from "react-dom/client"
import KorahCRM from "../KorahCRM.jsx"

import "./styles.css"

class ErrorBoundary extends React.Component {
  constructor(props) {
    super(props)
    this.state = { error: null }
  }

  static getDerivedStateFromError(error) {
    return { error }
  }

  render() {
    if (this.state.error) {
      return (
        <div className="boot-error">
          <strong>Erro ao carregar o CRM</strong>
          <span>{this.state.error.message}</span>
        </div>
      )
    }

    return this.props.children
  }
}

createRoot(document.getElementById("root")).render(
  <ErrorBoundary>
    <KorahCRM />
  </ErrorBoundary>
)
