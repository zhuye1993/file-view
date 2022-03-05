'use strict'

let _order = 1

export default function t_xml (S) {
  const openBracket = '<'
  const openBracketCC = '<'.charCodeAt(0)
  const closeBracket = '>'
  const closeBracketCC = '>'.charCodeAt(0)
  const minusCC = '-'.charCodeAt(0)
  const slashCC = '/'.charCodeAt(0)
  const exclamationCC = '!'.charCodeAt(0)
  const singleQuoteCC = '\''.charCodeAt(0)
  const doubleQuoteCC = '"'.charCodeAt(0)
  const questionMarkCC = '?'.charCodeAt(0)

  /**
   *    returns text until the first nonAlphebetic letter
   */
  const nameSpacer = '\r\n\t>/= '

  let pos = 0

  /**
   * Parsing a list of entries
   */
  function parseChildren () {
    const children = []
    while (S[pos]) {
      if (S.charCodeAt(pos) === openBracketCC) {
        if (S.charCodeAt(pos + 1) === slashCC) { // </
          // while (S[pos]!=='>') { pos++; }
          pos = S.indexOf(closeBracket, pos)
          return children
        } else if (S.charCodeAt(pos + 1) === exclamationCC) { // <! or <!--
          if (S.charCodeAt(pos + 2) === minusCC) {
            // comment support
            while (!(S.charCodeAt(pos) === closeBracketCC && S.charCodeAt(pos - 1) === minusCC &&
              S.charCodeAt(pos - 2) === minusCC && pos !== -1)) {
              pos = S.indexOf(closeBracket, pos + 1)
            }
            if (pos === -1) {
              pos = S.length
            }
          } else {
            // doctype support
            pos += 2
            for (; S.charCodeAt(pos) !== closeBracketCC; pos++) {}
          }
          pos++
          continue
        } else if (S.charCodeAt(pos + 1) === questionMarkCC) { // <?
          // XML header support
          pos = S.indexOf(closeBracket, pos)
          pos++
          continue
        }
        pos++
        let startNamePos = pos
        for (; nameSpacer.indexOf(S[pos]) === -1; pos++) {}
        const nodeTagName = S.slice(startNamePos, pos)

        // Parsing attributes
        let attrFound = false
        let nodeAttributes = {}
        for (; S.charCodeAt(pos) !== closeBracketCC; pos++) {
          const c = S.charCodeAt(pos)
          if ((c > 64 && c < 91) || (c > 96 && c < 123)) {
            startNamePos = pos
            for (; nameSpacer.indexOf(S[pos]) === -1; pos++) {}
            const name = S.slice(startNamePos, pos)
            // search beginning of the string
            let code = S.charCodeAt(pos)
            while (code !== singleQuoteCC && code !== doubleQuoteCC) {
              pos++
              code = S.charCodeAt(pos)
            }

            const startChar = S[pos]
            const startStringPos = ++pos
            pos = S.indexOf(startChar, startStringPos)
            const value = S.slice(startStringPos, pos)
            if (!attrFound) {
              nodeAttributes = {}
              attrFound = true
            }
            nodeAttributes[name] = value
          }
        }

        // Optional parsing of children
        let nodeChildren
        if (S.charCodeAt(pos - 1) !== slashCC) {
          pos++
          nodeChildren = parseChildren()
        }

        children.push({
          'children': nodeChildren,
          'tagName': nodeTagName,
          'attrs': nodeAttributes
        })
      } else {
        const startTextPos = pos
        pos = S.indexOf(openBracket, pos) - 1 // Skip characters until '<'
        if (pos === -2) {
          pos = S.length
        }
        const text = S.slice(startTextPos, pos + 1)
        if (text.trim().length > 0) {
          children.push(text)
        }
      }
      pos++
    }
    return children
  }

  _order = 1
  return simplefy(parseChildren())
}

function simplefy (children) {
  const node = {}

  if (children === undefined) {
    return {}
  }

  // Text node (e.g. <t>This is text.</t>)
  if (children.length === 1 && (typeof children[0] === 'string' || children[0] instanceof String)) {
    // eslint-disable-next-line no-new-wrappers
    return new String(children[0])
  }

  // map each object
  children.forEach(function (child) {
    if (!node[child.tagName]) {
      node[child.tagName] = []
    }

    if (typeof child === 'object') {
      const kids = simplefy(child.children)
      if (child.attrs) {
        kids.attrs = child.attrs
      }

      if (kids['attrs'] === undefined) {
        kids['attrs'] = {'order': _order}
      } else {
        kids['attrs']['order'] = _order
      }
      _order++
      node[child.tagName].push(kids)
    }
  })

  for (let i in node) {
    if (node[i].length === 1) {
      node[i] = node[i][0]
    }
  }

  return node
}
