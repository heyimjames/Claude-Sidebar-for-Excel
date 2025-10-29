import { useState, useRef, useEffect } from 'react';
import { ArrowDown24Regular, Settings24Regular } from '@fluentui/react-icons';
import Message from './Message';
import MessageInput from './MessageInput';
import ShortcutsHelp from './ShortcutsHelp';
import Settings from './Settings';
import ExcelContext from './ExcelContext';
import SuggestionChips from './SuggestionChips';
import ToolsMenu from './ToolsMenu';
import { ToolCallIndicator } from './ToolCallIndicator';
import { useClaudeChat } from '../hooks/useClaudeChat';
import { useKeyboardShortcuts } from '../hooks/useKeyboardShortcuts';
import { useScreenReaderAnnouncement } from '../hooks/useScreenReaderAnnouncement';
import { useExcelContext } from '../hooks/useExcelContext';
import { useSmartSuggestions } from '../hooks/useSmartSuggestions';
import type { ImageAttachment } from '../lib/types';
import '../styles/chat.css';

interface ChatInterfaceProps {
  apiKey: string;
}

export default function ChatInterface({ apiKey: initialApiKey }: ChatInterfaceProps) {
  const [apiKey, setApiKey] = useState(initialApiKey);
  const [input, setInput] = useState('');
  const [showShortcutsHelp, setShowShortcutsHelp] = useState(false);
  const [showSettings, setShowSettings] = useState(false);
  const [isAtBottom, setIsAtBottom] = useState(true);
  const [showScrollButton, setShowScrollButton] = useState(false);
  const messagesEndRef = useRef<HTMLDivElement>(null);
  const messageListRef = useRef<HTMLDivElement>(null);

  const { messages, isLoading, activeToolCalls, sendMessage, clearMessages, regenerateMessage, stopGeneration } = useClaudeChat(apiKey);
  const { announce } = useScreenReaderAnnouncement();
  const { context: excelContext, isLoading: isExcelLoading } = useExcelContext();
  const suggestions = useSmartSuggestions(excelContext);

  const scrollToBottom = (smooth = true) => {
    if (smooth) {
      messagesEndRef.current?.scrollIntoView({ behavior: 'smooth', block: 'end' });
    } else {
      messagesEndRef.current?.scrollIntoView({ behavior: 'auto', block: 'end' });
    }
  };

  // Check if user is at the bottom of the message list
  const checkIfAtBottom = () => {
    const messageList = messageListRef.current;
    if (!messageList) return true;

    const threshold = 100; // pixels from bottom
    const isBottom = messageList.scrollHeight - messageList.scrollTop - messageList.clientHeight < threshold;
    setIsAtBottom(isBottom);
    setShowScrollButton(!isBottom && messages.length > 0);
    return isBottom;
  };

  // Auto-scroll only if user is at the bottom
  useEffect(() => {
    if (isAtBottom) {
      scrollToBottom();
    }
  }, [messages, isAtBottom]);

  // Set up IntersectionObserver for the end marker
  useEffect(() => {
    const observer = new IntersectionObserver(
      (entries) => {
        const [entry] = entries;
        setIsAtBottom(entry.isIntersecting);
        setShowScrollButton(!entry.isIntersecting && messages.length > 0);
      },
      {
        root: messageListRef.current,
        threshold: 0.1,
      }
    );

    const endElement = messagesEndRef.current;
    if (endElement) {
      observer.observe(endElement);
    }

    return () => {
      if (endElement) {
        observer.unobserve(endElement);
      }
    };
  }, [messages.length]);

  // Handle scroll events for manual scrolling
  useEffect(() => {
    const messageList = messageListRef.current;
    if (!messageList) return;

    const handleScroll = () => {
      checkIfAtBottom();
    };

    messageList.addEventListener('scroll', handleScroll);
    return () => messageList.removeEventListener('scroll', handleScroll);
  }, [messages.length]);

  // Announce new messages to screen readers
  useEffect(() => {
    if (messages.length > 0) {
      const lastMessage = messages[messages.length - 1];
      if (lastMessage.role === 'assistant' && !lastMessage.isStreaming) {
        const content = typeof lastMessage.content === 'string' ? lastMessage.content : '';
        const preview = content.substring(0, 100);
        announce(`Claude responded: ${preview}${content.length > 100 ? '...' : ''}`);
      }
    }
  }, [messages, announce]);

  // Setup keyboard shortcuts
  useKeyboardShortcuts([
    {
      key: 'k',
      metaKey: true,
      callback: () => {
        // Focus on input
        document.querySelector<HTMLTextAreaElement>('#message-textarea')?.focus();
      },
    },
    {
      key: 'l',
      metaKey: true,
      callback: () => {
        clearMessages();
        announce('Chat history cleared');
      },
    },
    {
      key: '?',
      shiftKey: true,
      callback: () => {
        setShowShortcutsHelp(true);
      },
    },
  ]);

  const handleSendMessage = async (content: string, attachments?: ImageAttachment[]) => {
    setInput('');
    await sendMessage(content, excelContext, attachments);
  };

  const handleSuggestionClick = (prompt: string) => {
    setInput(prompt);
    // Auto-focus the input
    document.querySelector<HTMLTextAreaElement>('#message-textarea')?.focus();
  };

  const handleClearContext = async () => {
    try {
      await Excel.run(async (context) => {
        // Clear the current selection by selecting a single cell
        const range = context.workbook.worksheets.getActiveWorksheet().getRange('A1');
        range.select();
        await context.sync();
        announce('Excel context cleared');
      });
    } catch (error) {
      console.error('Clear context error:', error);
    }
  };

  const handleNewChat = () => {
    clearMessages();
    setInput('');
    announce('New chat started');
  };

  return (
    <div className="chat-interface" role="main" aria-label="Chat with Claude">
      <div className="chat-header" role="banner">
        <div className="header-icon">
          <svg width="24" height="24" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
            <path d="M12 2L2 7L12 12L22 7L12 2Z" stroke="currentColor" strokeWidth="2" strokeLinejoin="round" />
            <path d="M2 17L12 22L22 17" stroke="currentColor" strokeWidth="2" strokeLinejoin="round" />
            <path d="M2 12L12 17L22 12" stroke="currentColor" strokeWidth="2" strokeLinejoin="round" />
          </svg>
        </div>
        <div className="header-content">
          <h1 className="header-title">Claude</h1>
          <p className="header-subtitle">Excel AI Assistant</p>
        </div>
        <div className="header-actions">
          <ToolsMenu />
          <button
            className="icon-button"
            onClick={() => setShowSettings(true)}
            aria-label="Settings"
            title="Settings"
            type="button"
          >
            <Settings24Regular style={{ width: '14px', height: '14px' }} />
          </button>
          <button
            className="new-chat-button"
            onClick={handleNewChat}
            aria-label="Start new chat"
            title="Start new chat (⌘L)"
            type="button"
          >
            New Chat
          </button>
        </div>
      </div>

      <div
        ref={messageListRef}
        className="message-list"
        role="log"
        aria-label="Chat messages"
        aria-live="polite"
        aria-relevant="additions"
      >
        {messages.length === 0 ? (
          <div className="welcome-message">
            <h2>Welcome to Claude for Excel</h2>
            <p>I can help you understand and work with your Excel data. Try asking me to:</p>
            <ul>
              <li>Read and summarize your data</li>
              <li>Create charts and tables</li>
              <li>Apply formulas and formatting</li>
              <li>Analyze trends and patterns</li>
            </ul>
          </div>
        ) : (
          <>
            {messages.map((message) => (
              <Message key={message.id} message={message} onRegenerate={regenerateMessage} />
            ))}
            <ToolCallIndicator toolCalls={activeToolCalls} />
            {isLoading && (
              <div className="thinking-indicator">
                <span className="thinking-text shimmer">Thinking...</span>
              </div>
            )}
            <div ref={messagesEndRef} />
          </>
        )}
      </div>

      {/* Scroll to bottom button */}
      {showScrollButton && (
        <button
          className="scroll-to-bottom-button"
          onClick={() => scrollToBottom()}
          aria-label="Scroll to bottom"
          title="Scroll to bottom"
          type="button"
        >
          <ArrowDown24Regular />
        </button>
      )}

      <div className="chat-context-section">
        <ExcelContext context={excelContext} isLoading={isExcelLoading} />
        <SuggestionChips suggestions={suggestions} onSuggestionClick={handleSuggestionClick} />
      </div>

      <MessageInput
        value={input}
        onChange={setInput}
        onSend={handleSendMessage}
        onStop={stopGeneration}
        disabled={isLoading}
        isGenerating={isLoading}
        excelContext={excelContext}
        onClearContext={handleClearContext}
      />

      <ShortcutsHelp open={showShortcutsHelp} onClose={() => setShowShortcutsHelp(false)} />
      <Settings
        open={showSettings}
        onClose={() => setShowSettings(false)}
        apiKey={apiKey}
        onApiKeyChange={setApiKey}
      />
    </div>
  );
}
