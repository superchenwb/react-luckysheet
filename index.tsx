import ReactDOM from 'react-dom'
import React from 'react'
import { Luckysheet } from './src/index'

function App() {
  return (
    <div style={{width: 900, height: 600}}>
      <Luckysheet options={{lang: 'zh'}} />
    </div>
  )
}

ReactDOM.render(<App />, document.getElementById('root'))
