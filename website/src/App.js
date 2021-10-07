import { useState } from 'react';
import TextModifier from './components/TextModifier/TextModifier';

function App() {
  const [text, setText] = useState('');
  const [splitter, setSplitter] = useState('.');
  const handleChangeText = (ev) => {
    setText(ev.currentTarget.value);
  }
  const handleChangeSplitter = (ev) => {
    setSplitter(ev.currentTarget.value);
  }
  return (
    <div className="App">
      <TextModifier split={splitter}>{text}</TextModifier>
      <div className="panel">
        <textarea value={text} onChange={handleChangeText}/>
        <input classname="splitter" type="text" value={splitter} onChange={handleChangeSplitter} />
      </div>
    </div>
  );
}

export default App;
