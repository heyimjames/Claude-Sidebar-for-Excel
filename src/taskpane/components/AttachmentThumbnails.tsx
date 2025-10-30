import { Document24Regular, Image24Regular } from '@fluentui/react-icons';
import type { ImageAttachment } from '../lib/types';
import '../styles/attachment-thumbnails.css';

interface AttachmentThumbnailsProps {
  attachments: ImageAttachment[];
}

export default function AttachmentThumbnails({ attachments }: AttachmentThumbnailsProps) {
  if (!attachments || attachments.length === 0) return null;

  const maxVisible = 3;
  const visibleAttachments = attachments.slice(0, maxVisible);
  const remainingCount = attachments.length - maxVisible;

  return (
    <div className="attachment-thumbnails">
      {visibleAttachments.map((file, index) => (
        <div
          key={file.id}
          className="attachment-thumbnail"
          style={{
            '--stack-index': index,
            '--stack-total': visibleAttachments.length,
          } as React.CSSProperties}
          title={file.name || `Attachment ${index + 1}`}
        >
          {file.fileType === 'image' && file.previewUrl ? (
            <div className="thumbnail-content thumbnail-image">
              <img src={file.previewUrl} alt={file.name || 'Uploaded image'} />
            </div>
          ) : (
            <div className="thumbnail-content thumbnail-document">
              <Document24Regular className="thumbnail-icon" />
              <span className="thumbnail-label">
                {file.mediaType === 'application/pdf' ? 'PDF' : 'File'}
              </span>
            </div>
          )}
          {file.name && (
            <div className="thumbnail-filename">{file.name}</div>
          )}
        </div>
      ))}
      {remainingCount > 0 && (
        <div
          className="attachment-thumbnail attachment-more"
          style={{
            '--stack-index': maxVisible,
            '--stack-total': maxVisible + 1,
          } as React.CSSProperties}
          title={`${remainingCount} more file${remainingCount > 1 ? 's' : ''}`}
        >
          <div className="thumbnail-content thumbnail-more-indicator">
            <span className="more-count">+{remainingCount}</span>
          </div>
        </div>
      )}
    </div>
  );
}
