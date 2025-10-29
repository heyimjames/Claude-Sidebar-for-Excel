import ReactMarkdown from 'react-markdown';
import remarkGfm from 'remark-gfm';
import { Document24Regular } from '@fluentui/react-icons';
import type { ChatMessage } from '../lib/types';
import CodeBlock from './CodeBlock';
import MessageActions from './MessageActions';
import StreamingText from './StreamingText';
import CellReference, { detectCellReferences } from './CellReference';
import '../styles/message.css';

interface MessageProps {
  message: ChatMessage;
  onRegenerate?: (id: string) => void;
}

export default function Message({ message, onRegenerate }: MessageProps) {
  const isUser = message.role === 'user';

  // Extract text content from message (handle both string and array formats)
  const getTextContent = (): string => {
    if (typeof message.content === 'string') {
      return message.content;
    }
    // If content is an array, extract all text blocks
    const textBlocks = message.content.filter((block) => block.type === 'text');
    return textBlocks.map((block) => block.text).join('\n');
  };

  const textContent = getTextContent();

  return (
    <div
      className={`message ${isUser ? 'message-user' : 'message-assistant'}`}
      role="article"
      aria-label={`${isUser ? 'You' : 'Claude'} said`}
    >
      <div className="message-content">
        <MessageActions
          messageId={message.id}
          content={textContent}
          role={message.role}
          onRegenerate={onRegenerate}
        />

        {/* Display attached files for user messages */}
        {isUser && message.attachments && message.attachments.length > 0 && (
          <div className="message-images">
            {message.attachments.map((file) => (
              file.fileType === 'image' && file.previewUrl ? (
                <img
                  key={file.id}
                  src={file.previewUrl}
                  alt={file.name || 'Uploaded image'}
                  className="message-image"
                  title={file.name}
                />
              ) : (
                <div key={file.id} className="message-document" title={file.name}>
                  <Document24Regular />
                  {file.name && <span className="document-name">{file.name}</span>}
                </div>
              )
            ))}
          </div>
        )}

        <div className="message-text">
          <span className="sr-only">{isUser ? 'You:' : 'Claude:'}</span>
          {message.isAnimating && message.isStreaming ? (
            <StreamingText text={textContent} isComplete={!message.isStreaming} speed={50} />
          ) : (
            <ReactMarkdown
              remarkPlugins={[remarkGfm]}
              components={{
                code({ node, inline, className, children, ...props }) {
                  const match = /language-(\w+)/.exec(className || '');
                  const codeString = String(children).replace(/\n$/, '');

                  return !inline ? (
                    <CodeBlock code={codeString} language={match ? match[1] : 'text'} />
                  ) : (
                    <code className={className} {...props}>
                      {children}
                    </code>
                  );
                },
                // Custom text renderer to detect cell references
                p({ children }) {
                  if (typeof children === 'string') {
                    const { segments } = detectCellReferences(children);
                    return (
                      <p>
                        {segments.map((segment, index) =>
                          segment.type === 'cell' ? (
                            <CellReference key={index} reference={segment.content} />
                          ) : (
                            <span key={index}>{segment.content}</span>
                          )
                        )}
                      </p>
                    );
                  }
                  return <p>{children}</p>;
                },
              }}
            >
              {textContent}
            </ReactMarkdown>
          )}
        </div>
      </div>
    </div>
  );
}
