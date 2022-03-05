/**
 * Colorz (or Colz) is a Javascript "library" to help
 * in color conversion between the usual color-spaces
 * Hex - Rgb - Hsl / Hsv - Hsb
 *
 * It provides some helpers to output Canvas / CSS
 * color strings.
 *
 * by Carlos Cabo 2013
 * http://carloscabo.com
 *
 * Some formulas borrowed from Wikipedia or other authors.
 */

const round = Math.round

/*
 ==================================
 Color constructors
 ==================================
*/

export class Rgb {
  constructor (col) {
    this.r = col[0]
    this.g = col[1]
    this.b = col[2]
  }

  toString () {
    return `rgb(${this.r},${this.g},${this.b})`
  }
}

export class Rgba extends Rgb {
  constructor (col) {
    super(col)
    this.a = col[3]
  }

  toString () {
    return `rgba(${this.r},${this.g},${this.b},${this.a})`
  }
}

export class Hsl {
  constructor (col) {
    this.h = col[0]
    this.s = col[1]
    this.l = col[2]
  }

  toString () {
    return `hsl(${this.h},${this.s}%,${this.l}%)`
  }
}

export class Hsla extends Hsl {
  constructor (col) {
    super(col)
    this.a = col[3]
  }

  toString () {
    return `hsla(${this.h},${this.s}%,${this.l}%,${this.a})`
  }
}

/*
 ==================================
 Main Colz color object
 ==================================
*/
export class Color {
  constructor (r, g, b, a = 1.0) {
    // If args are not given in (r, g, b, [a]) format, convert
    if (typeof r === 'string') {
      let str = r
      // Add initial '#' if missing
      if (str.charAt(0) !== '#') { str = '#' + str }
      // If Hex in #fff format convert to #ffffff
      if (str.length < 7) {
        str = '#' + str[1] + str[1] + str[2] + str[2] + str[3] + str[3]
      }
      ([r, g, b] = hexToRgb(str))
    } else if (r instanceof Array) {
      a = r[3] || a
      b = r[2]
      g = r[1]
      r = r[0]
    }

    this.r = r
    this.g = g
    this.b = b
    this.a = a

    this.rgb = new Rgb([this.r, this.g, this.b])
    this.rgba = new Rgba([this.r, this.g, this.b, this.a])
    this.hex = rgbToHex(this.r, this.g, this.b)

    this.hsl = new Hsl(rgbToHsl(this.r, this.g, this.b))
    this.h = this.hsl.h
    this.s = this.hsl.s
    this.l = this.hsl.l
    this.hsla = new Hsla([this.h, this.s, this.l, this.a])
  }

  setHue (newHue) {
    this.h = newHue
    this.hsl.h = newHue
    this.hsla.h = newHue
    this.updateFromHsl()
  }

  setSat (newSat) {
    this.s = newSat
    this.hsl.s = newSat
    this.hsla.s = newSat
    this.updateFromHsl()
  }

  setLum (newLum) {
    this.l = newLum
    this.hsl.l = newLum
    this.hsla.l = newLum
    this.updateFromHsl()
  }

  setAlpha (newAlpha) {
    this.a = newAlpha
    this.hsla.a = newAlpha
    this.rgba.a = newAlpha
  }

  updateFromHsl () {
    // Updates Rgb
    this.rgb = null
    this.rgb = new Rgb(hslToRgb(this.h, this.s, this.l))

    this.r = this.rgb.r
    this.g = this.rgb.g
    this.b = this.rgb.b
    this.rgba.r = this.rgb.r
    this.rgba.g = this.rgb.g
    this.rgba.b = this.rgb.b

    // Updates Hex
    this.hex = null
    this.hex = rgbToHex([this.r, this.g, this.b])
  }
}

/*
 ==================================
 Public Methods
 ==================================
*/

export const randomColor = function () {
  const r = '#' + Math.random().toString(16).slice(2, 8)
  return new Color(r)
}

export const hexToRgb = function (hex) {
  const result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex)
  return result ? [
    parseInt(result[1], 16),
    parseInt(result[2], 16),
    parseInt(result[3], 16)
  ] : null
}

export const componentToHex = function (c) {
  const hex = c.toString(16)
  return hex.length === 1 ? '0' + hex : hex
}

// You can pass 3 numeric values or 1 Array
export const rgbToHex = function (r, g, b) {
  if (r instanceof Array) {
    b = r[2]
    g = r[1]
    r = r[0]
  }
  return '#' + componentToHex(r) + componentToHex(g) + componentToHex(b)
}

/**
 * Converts an RGB color value to HSL. Conversion formula
 * adapted from http://en.wikipedia.org/wiki/HSL_color_space.
 *
 * @param {Number} r The red color value
 * @param {Number} g The green color value
 * @param {Number} b The blue color value
 * @return {Array} The HSL representation
 */
export const rgbToHsl = function (r, g, b) {
  if (r instanceof Array) {
    b = r[2]
    g = r[1]
    r = r[0]
  }

  let h, s, l, d, max, min

  r /= 255
  g /= 255
  b /= 255

  max = Math.max(r, g, b)
  min = Math.min(r, g, b)
  l = (max + min) / 2

  if (max === min) {
    h = s = 0 // achromatic
  } else {
    d = max - min
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min)

    switch (max) {
      case r:
        h = (g - b) / d + (g < b ? 6 : 0)
        break
      case g:
        h = (b - r) / d + 2
        break
      case b:
        h = (r - g) / d + 4
        break
    }

    h /= 6
  }

  // CARLOS
  h = round(h * 360)
  s = round(s * 100)
  l = round(l * 100)

  return [h, s, l]
}

export const hue2rgb = function (p, q, t) {
  if (t < 0) { t += 1 }
  if (t > 1) { t -= 1 }
  if (t < 1 / 6) { return p + (q - p) * 6 * t }
  if (t < 1 / 2) { return q }
  if (t < 2 / 3) { return p + (q - p) * (2 / 3 - t) * 6 }
  return p
}

/**
 * Converts an HSL color value to RGB. Conversion formula
 * adapted from http://en.wikipedia.org/wiki/HSL_color_space.
 *
 * @param {Number} h The hue
 * @param {Number} s The saturation
 * @param {Number} l The lightness
 * @return {Array} The RGB representation
 */

export const hslToRgb = function (h, s, l) {
  if (h instanceof Array) {
    l = h[2]
    s = h[1]
    h = h[0]
  }
  h = h / 360
  s = s / 100
  l = l / 100

  let r, g, b, q, p

  if (s === 0) {
    r = g = b = l // achromatic
  } else {
    q = l < 0.5 ? l * (1 + s) : l + s - l * s
    p = 2 * l - q
    r = hue2rgb(p, q, h + 1 / 3)
    g = hue2rgb(p, q, h)
    b = hue2rgb(p, q, h - 1 / 3)
  }
  return [round(r * 255), round(g * 255), round(b * 255)]
}

/**
 * Converts an RGB color value to HSB / HSV. Conversion formula
 * adapted from http://en.wikipedia.org/wiki/HSV_color_space.
 *
 * @param {Number} r The red color value
 * @param {Number} g The green color value
 * @param {Number} b The blue color value
 * @return {Array} The HSB representation
 */
export const rgbToHsb = function (r, g, b) {
  let max, min, h, s, v, d

  r = r / 255
  g = g / 255
  b = b / 255

  max = Math.max(r, g, b)
  min = Math.min(r, g, b)
  v = max

  d = max - min
  s = max === 0 ? 0 : d / max

  if (max === min) {
    h = 0 // achromatic
  } else {
    switch (max) {
      case r:
        h = (g - b) / d + (g < b ? 6 : 0)
        break
      case g:
        h = (b - r) / d + 2
        break
      case b:
        h = (r - g) / d + 4
        break
    }
    h /= 6
  }

  // map top 360,100,100
  h = round(h * 360)
  s = round(s * 100)
  v = round(v * 100)

  return [h, s, v]
}

/**
 * Converts an HSB / HSV color value to RGB. Conversion formula
 * adapted from http://en.wikipedia.org/wiki/HSV_color_space.
 *
 * @param {Number} h The hue
 * @param {Number} s The saturation
 * @param {Number} v The value
 * @return {Array} The RGB representation
 */
export const hsbToRgb = function (h, s, v) {
  let r, g, b, i, f, p, q, t

  // h = h / 360;
  if (v === 0) { return [0, 0, 0] }

  s = s / 100
  v = v / 100
  h = h / 60

  i = Math.floor(h)
  f = h - i
  p = v * (1 - s)
  q = v * (1 - (s * f))
  t = v * (1 - (s * (1 - f)))

  if (i === 0) {
    r = v
    g = t
    b = p
  } else if (i === 1) {
    r = q
    g = v
    b = p
  } else if (i === 2) {
    r = p
    g = v
    b = t
  } else if (i === 3) {
    r = p
    g = q
    b = v
  } else if (i === 4) {
    r = t
    g = p
    b = v
  } else if (i === 5) {
    r = v
    g = p
    b = q
  }

  r = Math.floor(r * 255)
  g = Math.floor(g * 255)
  b = Math.floor(b * 255)

  return [r, g, b]
}

export const hsvToRgb = hsbToRgb // alias

/* Convert from Hsv */
export const hsbToHsl = function (h, s, b) {
  return rgbToHsl(hsbToRgb(h, s, b))
}

export const hsvToHsl = hsbToHsl // alias

/*
 ==================================
 Color Scheme Builder
 ==================================
*/
export class ColorScheme {
  constructor (colorVal, angleArray) {
    this.palette = []

    if (angleArray === undefined && colorVal instanceof Array) {
      // Asume you passing a color array ['#f00','#0f0'...]
      this.createFromColors(colorVal)
    } else {
      // Create scheme from color + hue angles
      this.createFromAngles(colorVal, angleArray)
    }
  }

  createFromColors (colorVal) {
    for (let i in colorVal) {
      if (colorVal.hasOwn(i)) {
        // console.log(colorVal[i]);
        this.palette.push(new Color(colorVal[i]))
      }
    }
    return this.palette
  }

  createFromAngles (colorVal, angleArray) {
    this.palette.push(new Color(colorVal))

    for (let i in angleArray) {
      if (angleArray.hasOwn(i)) {
        const tempHue = (this.palette[0].h + angleArray[i]) % 360
        this.palette.push(new Color(hslToRgb(tempHue, this.palette[0].s, this.palette[0].l)))
      }
    }
    return this.palette
  }

  /* Complementary colors constructors */
  static Compl (colorVal) {
    return new this(colorVal, [180])
  }

  /* Triad */
  static Triad (colorVal) {
    return new this(colorVal, [120, 240])
  }

  /* Tetrad */
  static Tetrad (colorVal) {
    return new this(colorVal, [60, 180, 240])
  }

  /* Analogous */
  static Analog (colorVal) {
    return new this(colorVal, [-45, 45])
  }

  /* Split complementary */
  static Split (colorVal) {
    return new this(colorVal, [150, 210])
  }

  /* Accented Analogous */
  static Accent (colorVal) {
    return new this(colorVal, [-45, 45, 180])
  }
}
