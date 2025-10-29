import { useState, useEffect, useRef } from 'react';
import '../styles/streaming-text.css';

interface StreamingTextProps {
  text: string;
  isComplete: boolean;
  speed?: number; // ms per word
}

export default function StreamingText({ text, isComplete, speed = 50 }: StreamingTextProps) {
  const [displayedWords, setDisplayedWords] = useState<string[]>([]);
  const words = text.split(/\s+/).filter(Boolean);
  const wordIndexRef = useRef(0);

  useEffect(() => {
    if (isComplete) {
      // Show all words immediately when complete
      setDisplayedWords(words);
      return;
    }

    // Reset when text changes significantly (new message)
    if (words.length < wordIndexRef.current) {
      wordIndexRef.current = 0;
      setDisplayedWords([]);
    }

    // Animate words incrementally
    if (wordIndexRef.current < words.length) {
      const timer = setTimeout(() => {
        wordIndexRef.current++;
        setDisplayedWords(words.slice(0, wordIndexRef.current));
      }, speed);

      return () => clearTimeout(timer);
    }
  }, [text, isComplete, words, speed]);

  return (
    <span className="streaming-text">
      {displayedWords.map((word, index) => (
        <span key={index} className="streaming-word">
          {word}{' '}
        </span>
      ))}
      {!isComplete && <span className="streaming-cursor" />}
    </span>
  );
}
