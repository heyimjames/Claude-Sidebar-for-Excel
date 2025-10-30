import { useRef, useEffect, useState } from 'react';
import { Dismiss24Regular, Document24Regular, Attach24Regular, Dismiss20Regular } from '@fluentui/react-icons';
import CommandPalette from './CommandPalette';
import { Command } from '../lib/commands';
import { isMac } from '../hooks/useKeyboardShortcuts';
import { useAnimatedPlaceholder } from '../hooks/useAnimatedPlaceholder';
import type { ImageAttachment } from '../lib/types';
import type { ExcelContext } from '../hooks/useExcelContext';
import '../styles/message-input.css';
import '../styles/image-upload.css';

interface MessageInputProps {
  value: string;
  onChange: (value: string) => void;
  onSend: (value: string, attachments?: ImageAttachment[]) => void;
  onStop?: () => void;
  disabled?: boolean;
  isGenerating?: boolean;
  excelContext?: ExcelContext;
  onClearContext?: () => void;
}

const MAX_FILES = 10;

export default function MessageInput({ value, onChange, onSend, onStop, disabled, isGenerating, excelContext, onClearContext }: MessageInputProps) {
  const textareaRef = useRef<HTMLTextAreaElement>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [showCommandPalette, setShowCommandPalette] = useState(false);
  const [commandQuery, setCommandQuery] = useState('');
  const [palettePosition, setPalettePosition] = useState({ top: 0, left: 0 });
  const [imageAttachments, setImageAttachments] = useState<ImageAttachment[]>([]);
  const [isDragging, setIsDragging] = useState(false);
  const { text: animatedPlaceholder, opacity: placeholderOpacity } = useAnimatedPlaceholder();

  useEffect(() => {
    if (textareaRef.current) {
      textareaRef.current.style.height = 'auto';
      textareaRef.current.style.height = `${Math.min(textareaRef.current.scrollHeight, 120)}px`;
    }
  }, [value]);

  // Handle paste events for images
  useEffect(() => {
    const handlePaste = async (e: ClipboardEvent) => {
      const items = e.clipboardData?.items;
      if (!items) return;

      for (let i = 0; i < items.length; i++) {
        const item = items[i];
        if (item.type.startsWith('image/')) {
          e.preventDefault();
          const file = item.getAsFile();
          if (file) {
            await handleImageFile(file);
          }
        }
      }
    };

    const textarea = textareaRef.current;
    if (textarea) {
      textarea.addEventListener('paste', handlePaste);
      return () => textarea.removeEventListener('paste', handlePaste);
    }
  }, []);

  const handleImageFile = async (file: File): Promise<void> => {
    // Check file limit
    if (imageAttachments.length >= MAX_FILES) {
      alert(`Maximum of ${MAX_FILES} files allowed per message.`);
      return;
    }

    // Validate file type
    const validImageTypes = ['image/jpeg', 'image/png', 'image/gif', 'image/webp'];
    const validDocumentTypes = ['application/pdf'];
    const allValidTypes = [...validImageTypes, ...validDocumentTypes];

    if (!allValidTypes.includes(file.type)) {
      alert('Please upload a valid file (JPEG, PNG, GIF, WebP, or PDF)');
      return;
    }

    // Validate file size (max 5MB for images, 10MB for PDFs)
    const maxSize = file.type === 'application/pdf' ? 10 * 1024 * 1024 : 5 * 1024 * 1024;
    if (file.size > maxSize) {
      const limit = file.type === 'application/pdf' ? '10MB' : '5MB';
      alert(`File is too large. Please use a file smaller than ${limit}.`);
      return;
    }

    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => {
        const dataUrl = reader.result as string;
        // Extract base64 data (remove "data:image/jpeg;base64," or "data:application/pdf;base64," prefix)
        const base64Data = dataUrl.split(',')[1];

        const fileType = validImageTypes.includes(file.type) ? 'image' : 'document';

        const attachment: ImageAttachment = {
          id: crypto.randomUUID(),
          type: 'base64',
          data: base64Data,
          mediaType: file.type as ImageAttachment['mediaType'],
          previewUrl: fileType === 'image' ? dataUrl : undefined,
          name: file.name,
          fileType,
        };

        setImageAttachments((prev) => [...prev, attachment]);
        resolve();
      };
      reader.onerror = () => reject(reader.error);
      reader.readAsDataURL(file);
    });
  };

  const handleFileSelect = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;

    // Check if adding these files would exceed the limit
    const remainingSlots = MAX_FILES - imageAttachments.length;
    if (files.length > remainingSlots) {
      alert(`You can only add ${remainingSlots} more file(s). Maximum ${MAX_FILES} files per message.`);
    }

    // Process files up to the limit
    for (let i = 0; i < Math.min(files.length, remainingSlots); i++) {
      await handleImageFile(files[i]);
    }

    // Reset input so the same file can be selected again
    if (fileInputRef.current) {
      fileInputRef.current.value = '';
    }
  };

  const handleRemoveImage = (id: string) => {
    setImageAttachments((prev) => prev.filter((img) => img.id !== id));
  };

  const handleImageButtonClick = () => {
    fileInputRef.current?.click();
  };

  const handleChange = (e: React.ChangeEvent<HTMLTextAreaElement>) => {
    const newValue = e.target.value;
    onChange(newValue);

    // Check if user typed '/'
    const lastChar = newValue[newValue.length - 1];
    const beforeLastChar = newValue[newValue.length - 2];

    if (lastChar === '/' && (!beforeLastChar || beforeLastChar === ' ' || beforeLastChar === '\n')) {
      setShowCommandPalette(true);
      setCommandQuery('');
      // Calculate position
      const rect = textareaRef.current?.getBoundingClientRect();
      if (rect) {
        setPalettePosition({ top: rect.top - 300, left: rect.left });
      }
    } else if (showCommandPalette) {
      // Extract command query
      const lastSlashIndex = newValue.lastIndexOf('/');
      const query = newValue.substring(lastSlashIndex + 1);

      // Close palette if user deletes the slash or adds space
      if (lastSlashIndex === -1 || query.includes(' ') || query.includes('\n')) {
        setShowCommandPalette(false);
      } else {
        setCommandQuery(query);
      }
    }
  };

  const handleCommandSelect = (command: Command) => {
    // Replace /command with template
    const lastSlashIndex = value.lastIndexOf('/');
    const newValue = value.substring(0, lastSlashIndex) + command.template;
    onChange(newValue);
    setShowCommandPalette(false);
    textareaRef.current?.focus();
  };

  const handleKeyDown = (e: React.KeyboardEvent) => {
    // Command palette is handling arrow keys
    if (showCommandPalette && (e.key === 'ArrowUp' || e.key === 'ArrowDown' || e.key === 'Enter')) {
      return;
    }

    // Cmd+Enter or Ctrl+Enter to send
    if (e.key === 'Enter' && (e.metaKey || e.ctrlKey)) {
      e.preventDefault();
      if (value.trim() || imageAttachments.length > 0) {
        handleSend();
      }
      return;
    }

    // Enter alone sends (existing behavior)
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      if (value.trim() || imageAttachments.length > 0) {
        handleSend();
      }
    }

    // Escape to clear or close palette
    if (e.key === 'Escape') {
      if (showCommandPalette) {
        setShowCommandPalette(false);
      } else {
        onChange('');
        setImageAttachments([]);
      }
    }
  };

  const handleSend = () => {
    if ((value.trim() || imageAttachments.length > 0) && !disabled) {
      onSend(value, imageAttachments.length > 0 ? imageAttachments : undefined);
      setImageAttachments([]);
    }
  };

  // Calculate context percentage (max 10,000 cells = 100%)
  const getContextPercentage = (): number => {
    if (!excelContext || !excelContext.hasData) return 0;
    const totalCells = excelContext.rowCount * excelContext.columnCount;
    const maxCells = 10000;
    return Math.min(Math.round((totalCells / maxCells) * 100), 100);
  };

  const contextPercentage = getContextPercentage();

  // Drag and drop handlers
  const handleDragEnter = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(true);
  };

  const handleDragLeave = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    // Only set dragging to false if leaving the wrapper entirely
    if (e.currentTarget === e.target) {
      setIsDragging(false);
    }
  };

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
  };

  const handleDrop = async (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(false);

    const files = Array.from(e.dataTransfer.files);
    const remainingSlots = MAX_FILES - imageAttachments.length;

    if (files.length > remainingSlots) {
      alert(`You can only add ${remainingSlots} more file(s). Maximum ${MAX_FILES} files per message.`);
    }

    // Process files up to the limit
    for (let i = 0; i < Math.min(files.length, remainingSlots); i++) {
      await handleImageFile(files[i]);
    }
  };

  return (
    <div className="message-input-container" role="form" aria-label="Send message to Claude">
      {showCommandPalette && (
        <CommandPalette
          query={commandQuery}
          onSelect={handleCommandSelect}
          onClose={() => setShowCommandPalette(false)}
          position={palettePosition}
        />
      )}

      {/* File previews (images and documents) */}
      {imageAttachments.length > 0 && (
        <div className="image-preview-container">
          {imageAttachments.map((file) => {
            // Extract file extension from filename
            const getFileExtension = (filename: string) => {
              const parts = filename.split('.');
              return parts.length > 1 ? `.${parts[parts.length - 1]}` : '';
            };

            const fileExtension = file.name ? getFileExtension(file.name) : '';

            return (
              <div
                key={file.id}
                className={`image-preview ${file.fileType === 'document' ? 'document-preview' : ''}`}
                title={file.name || 'Uploaded file'}
              >
                {file.fileType === 'image' && file.previewUrl ? (
                  <img src={file.previewUrl} alt={file.name || 'Uploaded image'} />
                ) : (
                  <div className="document-icon">
                    <Document24Regular />
                    {fileExtension && <div className="document-extension">{fileExtension}</div>}
                  </div>
                )}
                <button
                  className="image-remove-button"
                  onClick={() => handleRemoveImage(file.id)}
                  aria-label={`Remove ${file.name || 'file'}`}
                  type="button"
                  title="Remove file"
                >
                  <Dismiss24Regular />
                </button>
              </div>
            );
          })}
        </div>
      )}

      <div
        className={`message-input-wrapper ${isDragging ? 'dragging' : ''}`}
        onDragEnter={handleDragEnter}
        onDragLeave={handleDragLeave}
        onDragOver={handleDragOver}
        onDrop={handleDrop}
      >
        <label htmlFor="message-textarea" className="sr-only">
          Message to Claude
        </label>

        {/* Hidden file input */}
        <input
          ref={fileInputRef}
          type="file"
          accept="image/jpeg,image/png,image/gif,image/webp,application/pdf"
          multiple
          onChange={handleFileSelect}
          style={{ display: 'none' }}
          aria-label="Upload files (images or PDFs)"
        />

        <textarea
          id="message-textarea"
          ref={textareaRef}
          className="message-textarea"
          value={value}
          onChange={handleChange}
          onKeyDown={handleKeyDown}
          placeholder={animatedPlaceholder}
          disabled={disabled}
          rows={1}
          aria-label="Type your message"
          aria-describedby="input-hint"
          style={{ '--placeholder-opacity': placeholderOpacity } as React.CSSProperties}
        />

        <div className="input-controls">
          <div className="input-controls-left">
            {/* File attach button */}
            <button
              className="attach-button"
              onClick={handleImageButtonClick}
              disabled={disabled || imageAttachments.length >= MAX_FILES}
              aria-label={`Attach files (${imageAttachments.length}/${MAX_FILES})`}
              type="button"
              title={`Attach files (${imageAttachments.length}/${MAX_FILES})`}
            >
              <Attach24Regular />
              {imageAttachments.length > 0 && (
                <span className="attach-count">{imageAttachments.length}/{MAX_FILES}</span>
              )}
            </button>

            {/* Context indicator with circular progress */}
            <button
              className={`context-button ${contextPercentage > 0 ? 'has-context' : ''}`}
              onClick={onClearContext}
              aria-label={contextPercentage > 0 ? `Excel context: ${contextPercentage}% - Click to clear` : `No Excel context selected`}
              type="button"
              title={contextPercentage > 0 ? `Excel context: ${contextPercentage}% of max\nClick to clear context and disable cell prioritization` : `Select Excel cells to add context`}
            >
              <svg className="context-progress-ring" width="28" height="28" viewBox="0 0 28 28">
                {/* Background circle */}
                <circle
                  cx="14"
                  cy="14"
                  r="11"
                  fill="none"
                  stroke="var(--border)"
                  strokeWidth="2"
                />
                {/* Progress circle (only visible when there's context) */}
                {contextPercentage > 0 && (
                  <circle
                    cx="14"
                    cy="14"
                    r="11"
                    fill="none"
                    stroke="var(--accent)"
                    strokeWidth="2"
                    strokeDasharray={`${2 * Math.PI * 11}`}
                    strokeDashoffset={`${2 * Math.PI * 11 * (1 - contextPercentage / 100)}`}
                    strokeLinecap="round"
                    transform="rotate(-90 14 14)"
                  />
                )}
              </svg>
              {contextPercentage > 0 ? (
                <>
                  <span className="context-percentage">{contextPercentage}%</span>
                  <Dismiss20Regular className="context-dismiss-icon" />
                </>
              ) : null}
            </button>
          </div>

          <button
            className={`send-button ${isGenerating ? 'stop-button' : ''}`}
            onClick={isGenerating ? onStop : handleSend}
            disabled={!isGenerating && (disabled || (!value.trim() && imageAttachments.length === 0))}
            aria-label={isGenerating ? 'Stop generating' : (disabled ? 'Sending message...' : 'Send message (Enter)')}
            type="button"
          >
            {isGenerating ? (
              <svg width="16" height="16" viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg">
                <rect x="4" y="4" width="8" height="8" fill="currentColor" rx="1" />
              </svg>
            ) : (
              <svg width="16" height="16" viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg">
                <path
                  d="M8 12V4M8 4L5 7M8 4L11 7"
                  stroke="currentColor"
                  strokeWidth="2"
                  strokeLinecap="round"
                  strokeLinejoin="round"
                />
              </svg>
            )}
          </button>
        </div>

        {/* Drag overlay */}
        {isDragging && (
          <div className="drag-overlay">
            <Attach24Regular />
            <span>Drop files here</span>
          </div>
        )}
      </div>
      <div id="input-hint" className="input-hint" role="status" aria-live="polite">
        {isMac() ? '⌘' : 'Ctrl'}+Enter to send • Shift+Enter for new line • Type / for commands
        {imageAttachments.length > 0 && ' • Files will be analyzed and then discarded'}
      </div>
    </div>
  );
}
