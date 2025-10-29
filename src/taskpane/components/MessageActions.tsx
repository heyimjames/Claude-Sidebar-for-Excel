import { useState } from 'react';
import { Button, Tooltip } from '@fluentui/react-components';
import { Copy24Regular, ArrowCounterclockwise24Regular } from '@fluentui/react-icons';
import '../styles/message-actions.css';

interface MessageActionsProps {
  messageId: string;
  content: string;
  role: 'user' | 'assistant';
  onRegenerate?: (id: string) => void;
}

export default function MessageActions({ messageId, content, role, onRegenerate }: MessageActionsProps) {
  const [copied, setCopied] = useState(false);

  const handleCopy = async () => {
    await navigator.clipboard.writeText(content);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  return (
    <div className="message-actions">
      <Tooltip content={copied ? 'Copied!' : 'Copy message'} relationship="label">
        <Button
          appearance="subtle"
          size="small"
          icon={<Copy24Regular />}
          onClick={handleCopy}
          aria-label="Copy message"
        />
      </Tooltip>

      {role === 'assistant' && onRegenerate && (
        <Tooltip content="Regenerate response" relationship="label">
          <Button
            appearance="subtle"
            size="small"
            icon={<ArrowCounterclockwise24Regular />}
            onClick={() => onRegenerate(messageId)}
            aria-label="Regenerate response"
          />
        </Tooltip>
      )}
    </div>
  );
}
