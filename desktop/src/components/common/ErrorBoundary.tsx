import { type FC, type ReactNode, Component } from "react";
import { AlertTriangle } from "./IconMap";
import { useSettingsStore } from "../../store/settingsStore";
import { getAppText } from "../../i18n";

interface ErrorBoundaryProps {
  children: ReactNode;
}

interface ErrorBoundaryState {
  hasError: boolean;
  error: Error | null;
}

export class ErrorBoundary extends Component<ErrorBoundaryProps, ErrorBoundaryState> {
  constructor(props: ErrorBoundaryProps) {
    super(props);
    this.state = { hasError: false, error: null };
  }

  static getDerivedStateFromError(error: Error): ErrorBoundaryState {
    return { hasError: true, error };
  }

  render(): ReactNode {
    if (this.state.hasError) {
      const { language } = useSettingsStore.getState();
      const text = getAppText(language);
      return (
        <div className="error-boundary">
          <div className="error-boundary-icon"><AlertTriangle size={32} /></div>
          <div className="error-boundary-title">{text.common.errorOccurred}</div>
          <div className="error-boundary-message">{this.state.error?.message || text.common.unknownError}</div>
          <button
            className="error-boundary-retry"
            onClick={() => this.setState({ hasError: false, error: null })}
          >
            {text.common.retry}
          </button>
        </div>
      );
    }
    return this.props.children;
  }
}
