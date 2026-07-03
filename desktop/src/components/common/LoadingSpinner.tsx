import { type FC } from "react";

export const LoadingSpinner: FC = () => (
  <div className="loading-spinner">
    <div className="spinner" />
    <div className="loading-text">加载中...</div>
  </div>
);
