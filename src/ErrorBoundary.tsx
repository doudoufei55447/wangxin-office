import { Component, type ReactNode } from "react";

interface Props {
  children: ReactNode;
}
interface State {
  error: Error | null;
}

// 全局兜底：任何渲染期异常只降级当前视图并给出恢复入口，绝不白屏
export class ErrorBoundary extends Component<Props, State> {
  state: State = { error: null };

  static getDerivedStateFromError(error: Error): State {
    return { error };
  }

  componentDidCatch(error: Error, info: { componentStack?: string | null }) {
    console.error("[ErrorBoundary]", error, info.componentStack);
  }

  render() {
    if (this.state.error) {
      return (
        <div style={{ height: "100%", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", gap: 12, padding: 24 }}>
          <div style={{ fontSize: 15, fontWeight: 600 }}>界面出现异常，已安全拦截</div>
          <div style={{ fontSize: 12, color: "var(--muted-fg)", maxWidth: 560, wordBreak: "break-all", textAlign: "center" }}>
            {this.state.error.message}
          </div>
          <button className="btn primary" onClick={() => this.setState({ error: null })}>
            重试
          </button>
        </div>
      );
    }
    return this.props.children;
  }
}
