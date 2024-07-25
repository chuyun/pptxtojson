import tinycolor from 'tinycolor2'
import {v4 as uuidV4} from 'uuid'
import { getSchemeColorFromTheme } from './schemeColor.js'
import {
  applyShade,
  applyTint,
  applyLumOff,
  applyLumMod,
  applyHueMod,
  applySatMod,
  hslToRgb,
} from './color.js'

import {
  // base64ArrayBuffer,
  getTextByPathList,
  angleToDegrees,
  escapeHtml,
  getMimeType,
  toHex,
} from './utils.js'


export function getFillType(node) {
  let fillType = ''
  if (node['a:noFill']) fillType = 'NO_FILL'
  if (node['a:solidFill']) fillType = 'SOLID_FILL'
  if (node['a:gradFill']) fillType = 'GRADIENT_FILL'
  if (node['a:pattFill']) fillType = 'PATTERN_FILL'
  if (node['a:blipFill']) fillType = 'PIC_FILL'
  if (node['a:grpFill']) fillType = 'GROUP_FILL'

  return fillType
}

export async function getPicFill(type, node, warpObj) {
  const { uploadFn } = warpObj.options
  let img
  const rId = node['a:blip']['attrs']['r:embed']
  let imgPath
  if (type === 'slideBg' || type === 'slide') {
    imgPath = getTextByPathList(warpObj, ['slideResObj', rId, 'target'])
  }
  else if (type === 'slideLayoutBg') {
    imgPath = getTextByPathList(warpObj, ['layoutResObj', rId, 'target'])
  }
  else if (type === 'slideMasterBg') {
    imgPath = getTextByPathList(warpObj, ['masterResObj', rId, 'target'])
  }
  else if (type === 'themeBg') {
    imgPath = getTextByPathList(warpObj, ['themeResObj', rId, 'target'])
  }
  if (!imgPath) return imgPath

  img = getTextByPathList(warpObj, ['loaded-images', imgPath])
  if (!img) {
    imgPath = escapeHtml(imgPath)

    const imgExt = imgPath.split('.').pop()
    if (imgExt === 'xml') return undefined

    const blob = await warpObj['zip'].file(imgPath).async('blob')
    const imgMimeType = getMimeType(imgExt)
    if (uploadFn) {
      const name = imgPath.split('/').pop() || `${uuidV4()}.${imgExt.toLowerCase()}`
      const file = new File([blob], uuidV4() + name, { type: imgMimeType, lastModified: Date.now() })
      const res = await uploadFn([file])
      img = res?.[0]
    }
    else {
      // const imgArrayBuffer = await warpObj['zip'].file(imgPath).async('arraybuffer')
      // img = `data:${imgMimeType};base64,${base64ArrayBuffer(imgArrayBuffer)}`
      img = blob
    }
  }
  return img
}

export async function getBgPicFill(bgPr, sorce, warpObj) {
  const src = await getPicFill(sorce, bgPr['a:blipFill'], warpObj)
  const aBlipNode = bgPr['a:blipFill']['a:blip']
  const aBlipNodeTile = bgPr['a:blipFill']['a:tile']
  const aBlipNodeTileAttrs = aBlipNodeTile?.['attrs']

  const aphaModFixNode = getTextByPathList(aBlipNode, ['a:alphaModFix', 'attrs'])
  let opacity = 1
  if (aphaModFixNode && aphaModFixNode['amt'] && aphaModFixNode['amt'] !== '') {
    opacity = parseInt(aphaModFixNode['amt']) / 100000
  }
  // parseInt(getTextByPathList(solidFill, ['a:schemeClr', 'a:lumOff', 'attrs', 'val'])) / 100000
  // <a:tile>：定义平铺属性。
  // flip="none"：定义是否翻转图片（可选值：none, x, y, xy）。
  // algn="tl"：定义平铺的对齐方式（可选值：tl（左上）, t（上）, tr（右上）, l（左）, ctr（中心）, r（右）, bl（左下）, b（下）, br（右下））。
  // <a:tx>：定义水平平移量。
  // <a:ty>：定义垂直平移量。
  // <a:sx>：定义水平缩放比例。
  // <a:sy>：定义垂直缩放比例。

  let tile
  if (aBlipNodeTileAttrs) {
    tile = {
      flip: aBlipNodeTileAttrs['flip'],
      algn: aBlipNodeTileAttrs['algn'],
      tx: parseInt(aBlipNodeTileAttrs['tx']) / 100000,
      ty: parseInt(aBlipNodeTileAttrs['ty']) / 100000,
      sx: parseInt(aBlipNodeTileAttrs['sx']) / 100000,
      sy: parseInt(aBlipNodeTileAttrs['sy']) / 100000,
    }
  }
  return {
    src,
    opacity,
    // 平铺属性
    tile
  }
}

export function getBgGradientFill(bgPr, phClr, slideMasterContent, warpObj) {
  if (bgPr) {
    const grdFill = bgPr['a:gradFill']
    const gsLst = grdFill['a:gsLst']['a:gs']
    const color_ary = []

    for (let i = 0; i < gsLst.length; i++) {
      const lo_color = getSolidFill(gsLst[i], slideMasterContent['p:sldMaster']['p:clrMap']['attrs'], phClr, warpObj)
      const pos = getTextByPathList(gsLst[i], ['attrs', 'pos'])

      color_ary[i] = {
        pos: pos ? (pos / 1000 + '%') : '',
        color: `#${lo_color}`,
      }
    }
    const lin = grdFill['a:lin']
    let rot = 90
    if (lin) {
      rot = angleToDegrees(lin['attrs']['ang'])
      rot = rot + 90
    }

    return {
      rot,
      colors: color_ary.sort((a, b) => parseInt(a.pos) - parseInt(b.pos)),
    }
  }
  else if (phClr) return `#${phClr}`
  return null
}

export async function getSlideBackgroundFill(warpObj) {
  const slideContent = warpObj['slideContent']
  const slideLayoutContent = warpObj['slideLayoutContent']
  const slideMasterContent = warpObj['slideMasterContent']

  let bgPr = getTextByPathList(slideContent, ['p:sld', 'p:cSld', 'p:bg', 'p:bgPr'])

  let background = '#fff'
  let backgroundType = 'color'

  if (bgPr) {
    const bgFillTyp = getFillType(bgPr)

    if (bgFillTyp === 'SOLID_FILL') {
      const sldFill = bgPr['a:solidFill']
      let clrMapOvr
      const sldClrMapOvr = getTextByPathList(slideContent, ['p:sld', 'p:clrMapOvr', 'a:overrideClrMapping', 'attrs'])
      if (sldClrMapOvr) clrMapOvr = sldClrMapOvr
      else {
        const sldClrMapOvr = getTextByPathList(slideLayoutContent, ['p:sldLayout', 'p:clrMapOvr', 'a:overrideClrMapping', 'attrs'])
        if (sldClrMapOvr) clrMapOvr = sldClrMapOvr
        else clrMapOvr = getTextByPathList(slideMasterContent, ['p:sldMaster', 'p:clrMap', 'attrs'])
      }
      const sldBgClr = getSolidFill(sldFill, clrMapOvr, undefined, warpObj)
      background = `#${sldBgClr}`
    }
    else if (bgFillTyp === 'GRADIENT_FILL') {
      const gradientFill = getBgGradientFill(bgPr, undefined, slideMasterContent, warpObj)
      if (typeof gradientFill === 'string') {
        background = gradientFill
      }
      else if (gradientFill) {
        background = gradientFill
        backgroundType = 'gradient'
      }
    }
    else if (bgFillTyp === 'PIC_FILL') {
      background = await getBgPicFill(bgPr, 'slideBg', warpObj)
      backgroundType = 'image'
    }
  }
  else {
    bgPr = getTextByPathList(slideLayoutContent, ['p:sldLayout', 'p:cSld', 'p:bg', 'p:bgPr'])

    let clrMapOvr
    const sldClrMapOvr = getTextByPathList(slideLayoutContent, ['p:sldLayout', 'p:clrMapOvr', 'a:overrideClrMapping', 'attrs'])
    if (sldClrMapOvr) clrMapOvr = sldClrMapOvr
    else clrMapOvr = getTextByPathList(slideMasterContent, ['p:sldMaster', 'p:clrMap', 'attrs'])

    if (bgPr) {
      const bgFillTyp = getFillType(bgPr)
      if (bgFillTyp === 'SOLID_FILL') {
        const sldFill = bgPr['a:solidFill']
        const sldBgClr = getSolidFill(sldFill, clrMapOvr, undefined, warpObj)
        background = `#${sldBgClr}`
      }
      else if (bgFillTyp === 'GRADIENT_FILL') {
        const gradientFill = getBgGradientFill(bgPr, undefined, slideMasterContent, warpObj)
        if (typeof gradientFill === 'string') {
          background = gradientFill
        }
        else if (gradientFill) {
          background = gradientFill
          backgroundType = 'gradient'
        }
      }
      else if (bgFillTyp === 'PIC_FILL') {
        background = await getBgPicFill(bgPr, 'slideLayoutBg', warpObj)
        backgroundType = 'image'
      }
    }
    else {
      bgPr = getTextByPathList(slideMasterContent, ['p:sldMaster', 'p:cSld', 'p:bg', 'p:bgPr'])

      const clrMap = getTextByPathList(slideMasterContent, ['p:sldMaster', 'p:clrMap', 'attrs'])
      if (bgPr) {
        const bgFillTyp = getFillType(bgPr)
        if (bgFillTyp === 'SOLID_FILL') {
          const sldFill = bgPr['a:solidFill']
          const sldBgClr = getSolidFill(sldFill, clrMap, undefined, warpObj)
          background = `#${sldBgClr}`
        }
        else if (bgFillTyp === 'GRADIENT_FILL') {
          const gradientFill = getBgGradientFill(bgPr, undefined, slideMasterContent, warpObj)
          if (typeof gradientFill === 'string') {
            background = gradientFill
          }
          else if (gradientFill) {
            background = gradientFill
            backgroundType = 'gradient'
          }
        }
        else if (bgFillTyp === 'PIC_FILL') {
          background = await getBgPicFill(bgPr, 'slideMasterBg', warpObj)
          backgroundType = 'image'
        }
      }
    }
  }
  return {
    type: backgroundType,
    value: background,
  }
}

export function getShapeFill(node, isSvgMode, warpObj) {
  if (getTextByPathList(node, ['p:spPr', 'a:noFill'])) {
    return isSvgMode ? 'none' : ''
  }

  let fillColor
  if (!fillColor) {
    fillColor = getTextByPathList(node, ['p:spPr', 'a:solidFill', 'a:srgbClr', 'attrs', 'val'])
  }

  if (!fillColor) {
    // const schemeClr = 'a:' + getTextByPathList(node, ['p:spPr', 'a:solidFill', 'a:schemeClr', 'attrs', 'val'])
    const schemeClr = getTextByPathList(node, ['p:spPr', 'a:solidFill', 'a:schemeClr', 'attrs', 'val'])
    fillColor = getSchemeColorFromTheme(schemeClr, warpObj)
  }

  if (!fillColor) {
    const schemeClr = getTextByPathList(node, ['p:style', 'a:fillRef', 'a:schemeClr', 'attrs', 'val'])
    fillColor = getSchemeColorFromTheme(schemeClr, warpObj)
  }

  if (fillColor) {
    fillColor = `#${fillColor}`

    let lumMod = parseInt(getTextByPathList(node, ['p:spPr', 'a:solidFill', 'a:schemeClr', 'a:lumMod', 'attrs', 'val'])) / 100000
    let lumOff = parseInt(getTextByPathList(node, ['p:spPr', 'a:solidFill', 'a:schemeClr', 'a:lumOff', 'attrs', 'val'])) / 100000
    if (isNaN(lumMod)) lumMod = 1.0
    if (isNaN(lumOff)) lumOff = 0

    const color = tinycolor(fillColor).toHsl()
    const lum = color.l * lumMod + lumOff
    return tinycolor({ h: color.h, s: color.s, l: lum, a: color.a }).toHexString()
  }

  if (isSvgMode) return 'none'
  return fillColor
}

export function getSolidFill(solidFill, clrMap, phClr, warpObj) {
  if (!solidFill) return solidFill

  let color = 'fff'
  let clrNode

  if (solidFill['a:srgbClr']) {
    color = getTextByPathList(solidFill['a:srgbClr'], ['attrs', 'val'])
  }
  else if (solidFill['a:schemeClr']) {
    const schemeClr = getTextByPathList(solidFill['a:schemeClr'], ['attrs', 'val'])
    color = getSchemeColorFromTheme(schemeClr, warpObj) || '#ffffffff'

    let lumMod = parseInt(getTextByPathList(solidFill, ['a:schemeClr', 'a:lumMod', 'attrs', 'val'])) / 100000
    let lumOff = parseInt(getTextByPathList(solidFill, ['a:schemeClr', 'a:lumOff', 'attrs', 'val'])) / 100000
    if (isNaN(lumMod)) lumMod = 1.0
    if (isNaN(lumOff)) lumOff = 0

    color = tinycolor(color).toHsl()
    const lum = color.l * lumMod + lumOff
    return tinycolor({ h: color.h, s: color.s, l: lum, a: color.a }).toHex()
  }
  else if (solidFill['a:scrgbClr']) {
    clrNode = solidFill['a:scrgbClr']
    const defBultColorVals = clrNode['attrs']
    const red = (defBultColorVals['r'].indexOf('%') !== -1) ? defBultColorVals['r'].split('%').shift() : defBultColorVals['r']
    const green = (defBultColorVals['g'].indexOf('%') !== -1) ? defBultColorVals['g'].split('%').shift() : defBultColorVals['g']
    const blue = (defBultColorVals['b'].indexOf('%') !== -1) ? defBultColorVals['b'].split('%').shift() : defBultColorVals['b']
    color = toHex(255 * (Number(red) / 100)) + toHex(255 * (Number(green) / 100)) + toHex(255 * (Number(blue) / 100))
  }
  else if (solidFill['a:prstClr']) {
    clrNode = solidFill['a:prstClr']
    color = getTextByPathList(clrNode, ['attrs', 'val'])
  }
  else if (solidFill['a:hslClr']) {
    clrNode = solidFill['a:hslClr']
    const defBultColorVals = clrNode['attrs']
    const hue = Number(defBultColorVals['hue']) / 100000
    const sat = Number((defBultColorVals['sat'].indexOf('%') !== -1) ? defBultColorVals['sat'].split('%').shift() : defBultColorVals['sat']) / 100
    const lum = Number((defBultColorVals['lum'].indexOf('%') !== -1) ? defBultColorVals['lum'].split('%').shift() : defBultColorVals['lum']) / 100
    const hsl2rgb = hslToRgb(hue, sat, lum)
    color = toHex(hsl2rgb.r) + toHex(hsl2rgb.g) + toHex(hsl2rgb.b)
  }
  else if (solidFill['a:sysClr']) {
    clrNode = solidFill['a:sysClr']
    const sysClr = getTextByPathList(clrNode, ['attrs', 'lastClr'])
    if (sysClr) color = sysClr
  }

  let isAlpha = false
  const alpha = parseInt(getTextByPathList(clrNode, ['a:alpha', 'attrs', 'val'])) / 100000
  if (!isNaN(alpha)) {
    const al_color = tinycolor(color)
    al_color.setAlpha(alpha)
    color = al_color.toHex8()
    isAlpha = true
  }

  const hueMod = parseInt(getTextByPathList(clrNode, ['a:hueMod', 'attrs', 'val'])) / 100000
  if (!isNaN(hueMod)) {
    color = applyHueMod(color, hueMod, isAlpha)
  }
  const lumMod = parseInt(getTextByPathList(clrNode, ['a:lumMod', 'attrs', 'val'])) / 100000
  if (!isNaN(lumMod)) {
    color = applyLumMod(color, lumMod, isAlpha)
  }
  const lumOff = parseInt(getTextByPathList(clrNode, ['a:lumOff', 'attrs', 'val'])) / 100000
  if (!isNaN(lumOff)) {
    color = applyLumOff(color, lumOff, isAlpha)
  }
  const satMod = parseInt(getTextByPathList(clrNode, ['a:satMod', 'attrs', 'val'])) / 100000
  if (!isNaN(satMod)) {
    color = applySatMod(color, satMod, isAlpha)
  }
  const shade = parseInt(getTextByPathList(clrNode, ['a:shade', 'attrs', 'val'])) / 100000
  if (!isNaN(shade)) {
    color = applyShade(color, shade, isAlpha)
  }
  const tint = parseInt(getTextByPathList(clrNode, ['a:tint', 'attrs', 'val'])) / 100000
  if (!isNaN(tint)) {
    color = applyTint(color, tint, isAlpha)
  }

  return color
}
