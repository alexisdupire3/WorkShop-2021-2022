import './TextModifier.css';
import { useEffect, useRef, useState } from 'react';

const TextModifier = ({children, split}) => {
  const [words, setWords] = useState([]);
  const [index, setIndex] = useState(0);

  const indexRef = useRef(index);
  indexRef.current = index;

  useEffect(()=>{
    const newWords = children.split(split);
    setWords(newWords);
  },[children, split]);

  const endWords = ()=> {
    setIndex(words.length -1 );
  };

  useEffect(()=>{
    setIndex(0);
  },[split]);

  useEffect(()=>{
    document.addEventListener('keydown', function(event) {
        if(event.keyCode === 37) {
            setIndex(indexRef.current-1 === -1 ? 0:indexRef.current-1)
        }
        else if(event.keyCode === 39) {
            setIndex(indexRef.current+1 === words.length ? indexRef.current : indexRef.current+1)
        }
    });
  },[]);
  return (
    <p className="text_modifier">
        { words.map((word,key) =><span key={key} className={index===key ? 'display' : ''} onClick={()=>{setIndex(key)}}>{word}{split}</span>)}
    </p>
  );
}

export default TextModifier;
