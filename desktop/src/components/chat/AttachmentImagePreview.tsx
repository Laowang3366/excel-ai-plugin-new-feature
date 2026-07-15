import React, { useEffect, useMemo, useState } from "react";
import { createPortal } from "react-dom";
import type { FileAttachment } from "../../electronApi";
import { ipcApi } from "../../services/ipcApi";
import { buildImageDataUri, formatAttachmentSize } from "../../utils/attachmentPreview";
import { ExternalLink, Image, X } from "../common/IconMap";

type PreviewStatus = "loading" | "ready" | "error";

interface CachedPreview {
  src: string;
  size?: number;
}

interface PreviewState {
  status: PreviewStatus;
  src?: string;
  size?: number;
}

interface AttachmentImagePreviewProps {
  attachment: FileAttachment;
  variant: "composer" | "message";
  onRemove?: () => void;
}

const previewCache = new Map<string, CachedPreview>();

export const AttachmentImagePreview: React.FC<AttachmentImagePreviewProps> = React.memo(
  ({ attachment, variant, onRemove }) => {
    const [preview, setPreview] = useState<PreviewState>({ status: "loading" });
    const [lightboxOpen, setLightboxOpen] = useState(false);
    const cacheKey = attachment.filePath;

    useEffect(() => {
      let cancelled = false;
      const cached = previewCache.get(cacheKey);
      if (cached) {
        setPreview({ status: "ready", src: cached.src, size: cached.size ?? attachment.size });
        return () => {
          cancelled = true;
        };
      }

      setPreview({ status: "loading", size: attachment.size });
      ipcApi.file
        .readAsBase64(attachment.filePath)
        .then((result) => {
          if (cancelled) return;
          if (result.error || !result.data) {
            setPreview({ status: "error", size: attachment.size });
            return;
          }
          const src = buildImageDataUri(
            result.data,
            result.mimeType,
            result.fileName || attachment.fileName || attachment.filePath,
          );
          const next = { src, size: result.size ?? attachment.size };
          previewCache.set(cacheKey, next);
          setPreview({ status: "ready", ...next });
        })
        .catch(() => {
          if (!cancelled) setPreview({ status: "error", size: attachment.size });
        });

      return () => {
        cancelled = true;
      };
    }, [attachment.fileName, attachment.filePath, attachment.size, cacheKey]);

    const sizeLabel = useMemo(
      () => formatAttachmentSize(preview.size ?? attachment.size),
      [attachment.size, preview.size],
    );
    const canPreview = preview.status === "ready" && Boolean(preview.src);

    const openLightbox = () => {
      if (canPreview) setLightboxOpen(true);
    };

    const openOriginal = (event: React.MouseEvent) => {
      event.stopPropagation();
      void ipcApi.file.openFile(attachment.filePath);
    };

    const lightbox =
      lightboxOpen && canPreview && typeof document !== "undefined"
        ? createPortal(
            <div className="image-preview-lightbox" onClick={() => setLightboxOpen(false)}>
              <div className="image-preview-panel" onClick={(event) => event.stopPropagation()}>
                <div className="image-preview-toolbar">
                  <span className="image-preview-title" title={attachment.fileName}>
                    {attachment.fileName}
                  </span>
                  <div className="image-preview-toolbar-actions">
                    <button
                      type="button"
                      className="image-preview-icon-btn"
                      onClick={openOriginal}
                      title="打开原文件"
                      aria-label="打开原文件"
                    >
                      <ExternalLink size={17} />
                    </button>
                    <button
                      type="button"
                      className="image-preview-icon-btn"
                      onClick={() => setLightboxOpen(false)}
                      title="关闭预览"
                      aria-label="关闭预览"
                    >
                      <X size={18} />
                    </button>
                  </div>
                </div>
                <div className="image-preview-canvas">
                  <img src={preview.src} alt={attachment.fileName} draggable={false} />
                </div>
              </div>
            </div>,
            document.body,
          )
        : null;

    return (
      <div className={`image-attachment-preview ${variant}-image-preview ${preview.status}`}>
        <button
          type="button"
          className="image-attachment-thumb"
          onClick={openLightbox}
          disabled={!canPreview}
          title={canPreview ? "预览图片" : attachment.fileName}
          aria-label={canPreview ? `预览图片 ${attachment.fileName}` : attachment.fileName}
        >
          {canPreview ? (
            <img src={preview.src} alt={attachment.fileName} draggable={false} />
          ) : (
            <span className="image-attachment-placeholder">
              <Image size={18} />
            </span>
          )}
        </button>
        <div className="image-attachment-info">
          <span className="image-attachment-name" title={attachment.fileName}>
            {attachment.fileName}
          </span>
          {sizeLabel && <span className="image-attachment-size">{sizeLabel}</span>}
        </div>
        {onRemove && (
          <button
            type="button"
            className="image-attachment-remove"
            onClick={(event) => {
              event.stopPropagation();
              onRemove();
            }}
            title="移除附件"
            aria-label={`移除附件 ${attachment.fileName}`}
          >
            <X size={12} />
          </button>
        )}
        {lightbox}
      </div>
    );
  },
);
