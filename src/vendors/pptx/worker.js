'use strict'

import processPptx from './process_pptx'

processPptx(
  func => { self.onmessage = e => func(e.data) },
  msg => self.postMessage(msg)
)
