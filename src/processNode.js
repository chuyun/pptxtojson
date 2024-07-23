import {
  angleToDegrees,
  base64ArrayBuffer, escapeHtml,
  extractFileExtension,
  getMimeType,
  getTextByPathList,
  isVideoLink
} from './utils.js'
import {getPosition, getSize} from './position.js'
import {genTextBody} from './text.js'
import {getBorder} from './border.js'
import {getShapeFill, getSolidFill} from './fill.js'
import {getShadow} from './shadow.js'
import {getVerticalAlign} from './align.js'
import {getCustomShapePath} from './shape.js'
import {readXmlFile} from './readXmlFile.js'
import {getChartInfo} from './chart.js'

export const media_cache_dir = 'extracted_media_cache'

export async function processNodesInSlide(nodeKey, nodeValue, warpObj, source) {
  let json

  switch (nodeKey) {
    case 'p:sp': // Shape, Text
      json = processSpNode(nodeValue, warpObj, source)
      break
    case 'p:cxnSp': // Shape, Text
      json = processCxnSpNode(nodeValue, warpObj, source)
      break
    case 'p:pic': // Image, Video, Audio
      json = processPicNode(nodeValue, warpObj, source)
      break
    case 'p:graphicFrame': // Chart, Diagram, Table
      json = await processGraphicFrameNode(nodeValue, warpObj, source)
      break
    case 'p:grpSp':
      json = await processGroupSpNode(nodeValue, warpObj, source)
      break
    case 'mc:AlternateContent':
      json = await processGroupSpNode(getTextByPathList(nodeValue, ['mc:Fallback']), warpObj, source)
      break
    default:
  }

  return json
}

export async function processGroupSpNode(node, warpObj, source) {
  const xfrmNode = getTextByPathList(node, ['p:grpSpPr', 'a:xfrm'])
  if (!xfrmNode) return null

  const x = parseInt(xfrmNode['a:off']['attrs']['x']) * warpObj.options.slideFactor
  const y = parseInt(xfrmNode['a:off']['attrs']['y']) * warpObj.options.slideFactor
  // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.childoffset?view=openxml-2.8.1
  const chx = parseInt(xfrmNode['a:chOff']['attrs']['x']) * warpObj.options.slideFactor
  const chy = parseInt(xfrmNode['a:chOff']['attrs']['y']) * warpObj.options.slideFactor
  const cx = parseInt(xfrmNode['a:ext']['attrs']['cx']) * warpObj.options.slideFactor
  const cy = parseInt(xfrmNode['a:ext']['attrs']['cy']) * warpObj.options.slideFactor
  // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.childextents?view=openxml-2.8.1
  const chcx = parseInt(xfrmNode['a:chExt']['attrs']['cx']) * warpObj.options.slideFactor
  const chcy = parseInt(xfrmNode['a:chExt']['attrs']['cy']) * warpObj.options.slideFactor
  // children coordinate
  const ws = cx / chcx
  const hs = cy / chcy

  const elements = []
  for (const nodeKey in node) {
    if (node[nodeKey].constructor === Array) {
      for (const item of node[nodeKey]) {
        const ret = await processNodesInSlide(nodeKey, item, warpObj, source)
        if (ret) elements.push(ret)
      }
    }
    else {
      const ret = await processNodesInSlide(nodeKey, node[nodeKey], warpObj, source)
      if (ret) elements.push(ret)
    }
  }

  return {
    type: 'group',
    top: parseFloat(y.toFixed(2)),
    left: parseFloat(x.toFixed(2)),
    width: parseFloat(cx.toFixed(2)),
    height: parseFloat(cy.toFixed(2)),
    elements: elements.map(element => ({
      ...element,
      left: parseFloat(((element.left - chx) * ws).toFixed(2)),
      top: parseFloat(((element.top - chy) * hs).toFixed(2)),
      width: parseFloat((element.width * ws).toFixed(2)),
      height: parseFloat((element.height * hs).toFixed(2)),
    }))
  }
}

export function processSpNode(node, warpObj, source) {
  const name = getTextByPathList(node, ['p:nvSpPr', 'p:cNvPr', 'attrs', 'name'])
  const idx = getTextByPathList(node, ['p:nvSpPr', 'p:nvPr', 'p:ph', 'attrs', 'idx'])
  let type = getTextByPathList(node, ['p:nvSpPr', 'p:nvPr', 'p:ph', 'attrs', 'type'])

  let slideLayoutSpNode, slideMasterSpNode

  if (type) {
    if (idx) {
      slideLayoutSpNode = warpObj['slideLayoutTables']['typeTable'][type]
      slideMasterSpNode = warpObj['slideMasterTables']['typeTable'][type]
    }
    else {
      slideLayoutSpNode = warpObj['slideLayoutTables']['typeTable'][type]
      slideMasterSpNode = warpObj['slideMasterTables']['typeTable'][type]
    }
  }
  else if (idx) {
    slideLayoutSpNode = warpObj['slideLayoutTables']['idxTable'][idx]
    slideMasterSpNode = warpObj['slideMasterTables']['idxTable'][idx]
  }

  if (!type) {
    const txBoxVal = getTextByPathList(node, ['p:nvSpPr', 'p:cNvSpPr', 'attrs', 'txBox'])
    if (txBoxVal === '1') type = 'text'
  }
  if (!type) type = getTextByPathList(slideLayoutSpNode, ['p:nvSpPr', 'p:nvPr', 'p:ph', 'attrs', 'type'])
  if (!type) type = getTextByPathList(slideMasterSpNode, ['p:nvSpPr', 'p:nvPr', 'p:ph', 'attrs', 'type'])

  if (!type) {
    if (source === 'diagramBg') type = 'diagram'
    else type = 'obj'
  }

  return genShape(node, slideLayoutSpNode, slideMasterSpNode, name, type, warpObj)
}

export function processCxnSpNode(node, warpObj) {
  const name = node['p:nvCxnSpPr']['p:cNvPr']['attrs']['name']
  const type = (node['p:nvCxnSpPr']['p:nvPr']['p:ph'] === undefined) ? undefined : node['p:nvSpPr']['p:nvPr']['p:ph']['attrs']['type']

  return genShape(node, undefined, undefined, name, type, warpObj)
}

export function genShape(node, slideLayoutSpNode, slideMasterSpNode, name, type, warpObj) {
  const xfrmList = ['p:spPr', 'a:xfrm']
  const slideXfrmNode = getTextByPathList(node, xfrmList)
  const slideLayoutXfrmNode = getTextByPathList(slideLayoutSpNode, xfrmList)
  const slideMasterXfrmNode = getTextByPathList(slideMasterSpNode, xfrmList)

  const shapType = getTextByPathList(node, ['p:spPr', 'a:prstGeom', 'attrs', 'prst'])
  const custShapType = getTextByPathList(node, ['p:spPr', 'a:custGeom'])

  const { top, left } = getPosition(slideXfrmNode, slideLayoutXfrmNode, slideMasterXfrmNode, warpObj.options.slideFactor)
  const { width, height } = getSize(slideXfrmNode, slideLayoutXfrmNode, slideMasterXfrmNode, warpObj.options.slideFactor)

  const isFlipV = getTextByPathList(slideXfrmNode, ['attrs', 'flipV']) === '1'
  const isFlipH = getTextByPathList(slideXfrmNode, ['attrs', 'flipH']) === '1'

  const rotate = angleToDegrees(getTextByPathList(slideXfrmNode, ['attrs', 'rot']))

  const txtXframeNode = getTextByPathList(node, ['p:txXfrm'])
  let txtRotate
  if (txtXframeNode) {
    const txtXframeRot = getTextByPathList(txtXframeNode, ['attrs', 'rot'])
    if (txtXframeRot) txtRotate = angleToDegrees(txtXframeRot) + 90
  }
  else txtRotate = rotate

  let content = ''
  if (node['p:txBody']) content = genTextBody(node['p:txBody'], node, slideLayoutSpNode, type, warpObj)

  const { borderColor, borderWidth, borderType, strokeDasharray } = getBorder(node, type, warpObj)
  const fillColor = getShapeFill(node, undefined, warpObj) || ''

  let shadow
  const outerShdwNode = getTextByPathList(node, ['p:spPr', 'a:effectLst', 'a:outerShdw'])
  if (outerShdwNode) shadow = getShadow(outerShdwNode, warpObj)

  const vAlign = getVerticalAlign(node, slideLayoutSpNode, slideMasterSpNode, type)
  const isVertical = getTextByPathList(node, ['p:txBody', 'a:bodyPr', 'attrs', 'vert']) === 'eaVert'

  const data = {
    left,
    top,
    width,
    height,
    borderColor,
    borderWidth,
    borderType,
    borderStrokeDasharray: strokeDasharray,
    fillColor,
    content,
    isFlipV,
    isFlipH,
    rotate,
    vAlign,
    name,
  }

  if (shadow) data.shadow = shadow

  if (custShapType && type !== 'diagram') {
    const ext = getTextByPathList(slideXfrmNode, ['a:ext', 'attrs'])
    const w = parseInt(ext['cx']) * warpObj.options.slideFactor
    const h = parseInt(ext['cy']) * warpObj.options.slideFactor
    const d = getCustomShapePath(custShapType, w, h)

    return {
      ...data,
      type: 'shape',
      shapType: 'custom',
      path: d,
    }
  }
  if (shapType && (type === 'obj' || !type)) {
    return {
      ...data,
      type: 'shape',
      shapType,
    }
  }
  return {
    ...data,
    type: 'text',
    isVertical,
    rotate: txtRotate,
  }
}

export async function processPicNode(node, warpObj, source) {
  let resObj
  if (source === 'slideMasterBg') resObj = warpObj['masterResObj']
  else if (source === 'slideLayoutBg') resObj = warpObj['layoutResObj']
  else resObj = warpObj['slideResObj']

  const rid = node['p:blipFill']['a:blip']['attrs']['r:embed']
  const imgName = resObj[rid]['target']
  const imgFileExt = extractFileExtension(imgName).toLowerCase()
  const zip = warpObj['zip']
  const imgArrayBuffer = await zip.file(imgName).async('arraybuffer')
  const xfrmNode = node['p:spPr']['a:xfrm']

  const mimeType = getMimeType(imgFileExt)
  const { top, left } = getPosition(xfrmNode, undefined, undefined, warpObj.options.slideFactor)
  const { width, height } = getSize(xfrmNode, undefined, undefined, warpObj.options.slideFactor)
  const src = `data:${mimeType};base64,${base64ArrayBuffer(imgArrayBuffer)}`

  let srcPath = media_cache_dir + imgName.replace('ppt/media', '')

  const isFlipV = getTextByPathList(xfrmNode, ['attrs', 'flipV']) === '1'
  const isFlipH = getTextByPathList(xfrmNode, ['attrs', 'flipH']) === '1'

  let rotate = 0
  const rotateNode = getTextByPathList(node, ['p:spPr', 'a:xfrm', 'attrs', 'rot'])
  if (rotateNode) rotate = angleToDegrees(rotateNode)

  const videoNode = getTextByPathList(node, ['p:nvPicPr', 'p:nvPr', 'a:videoFile'])
  let videoRid, videoFile, videoFileExt, videoMimeType, uInt8ArrayVideo, videoBlob
  let isVdeoLink = false

  if (videoNode) {
    videoRid = videoNode['attrs']['r:link']
    videoFile = resObj[videoRid]['target']
    if (isVideoLink(videoFile)) {
      videoFile = escapeHtml(videoFile)
      isVdeoLink = true
    }
    else {
      videoFileExt = extractFileExtension(videoFile).toLowerCase()
      if (videoFileExt === 'mp4' || videoFileExt === 'webm' || videoFileExt === 'ogg') {
        srcPath = media_cache_dir + videoFile.replace('ppt/media', '')
        // TODO
        try {
          uInt8ArrayVideo = await zip.file(videoFile).async('arraybuffer')
          videoMimeType = getMimeType(videoFileExt)
          videoBlob = URL.createObjectURL(new Blob([uInt8ArrayVideo], {
            type: videoMimeType
          }))
        }
        catch (e) {
          console.log('e', e)
        }
      }
    }
  }

  const audioNode = getTextByPathList(node, ['p:nvPicPr', 'p:nvPr', 'a:audioFile'])
  let audioRid, audioFile, audioFileExt, uInt8ArrayAudio, audioBlob
  if (audioNode) {
    audioRid = audioNode['attrs']['r:link']
    audioFile = resObj[audioRid]['target']
    audioFileExt = extractFileExtension(audioFile).toLowerCase()
    if (audioFileExt === 'mp3' || audioFileExt === 'wav' || audioFileExt === 'ogg') {
      uInt8ArrayAudio = await zip.file(audioFile).async('arraybuffer')
      audioBlob = URL.createObjectURL(new Blob([uInt8ArrayAudio]))
    }
  }

  if (videoNode && !isVdeoLink) {
    return {
      type: 'video',
      top,
      left,
      width,
      height,
      rotate,
      blob: videoBlob,
      realPath: srcPath,
    }
  }
  if (videoNode && isVdeoLink) {
    return {
      type: 'video',
      top,
      left,
      width,
      height,
      rotate,
      src: videoFile,
      realPath: srcPath,
    }
  }
  if (audioNode) {
    return {
      type: 'audio',
      top,
      left,
      width,
      height,
      rotate,
      blob: audioBlob,
      realPath: srcPath,
    }
  }
  return {
    type: 'image',
    top,
    left,
    width,
    height,
    rotate,
    src,
    realPath: srcPath,
    isFlipV,
    isFlipH
  }
}

export async function processGraphicFrameNode(node, warpObj, source) {
  const graphicTypeUri = getTextByPathList(node, ['a:graphic', 'a:graphicData', 'attrs', 'uri'])

  let result
  switch (graphicTypeUri) {
    case 'http://schemas.openxmlformats.org/drawingml/2006/table':
      result = genTable(node, warpObj)
      break
    case 'http://schemas.openxmlformats.org/drawingml/2006/chart':
      result = await genChart(node, warpObj)
      break
    case 'http://schemas.openxmlformats.org/drawingml/2006/diagram':
      result = genDiagram(node, warpObj)
      break
    case 'http://schemas.openxmlformats.org/presentationml/2006/ole':
      let oleObjNode = getTextByPathList(node, ['a:graphic', 'a:graphicData', 'mc:AlternateContent', 'mc:Fallback', 'p:oleObj'])
      if (!oleObjNode) oleObjNode = getTextByPathList(node, ['a:graphic', 'a:graphicData', 'p:oleObj'])
      else processGroupSpNode(oleObjNode, warpObj, source)
      break
    default:
  }
  return result
}

export function genTable(node, warpObj) {
  const tableNode = getTextByPathList(node, ['a:graphic', 'a:graphicData', 'a:tbl'])
  const xfrmNode = getTextByPathList(node, ['p:xfrm'])
  const { top, left } = getPosition(xfrmNode, undefined, undefined, warpObj.options.slideFactor)
  const { width, height } = getSize(xfrmNode, undefined, undefined, warpObj.options.slideFactor)

  const getTblPr = getTextByPathList(node, ['a:graphic', 'a:graphicData', 'a:tbl', 'a:tblPr'])

  let thisTblStyle
  const tbleStyleId = getTblPr['a:tableStyleId']
  if (tbleStyleId) {
    const tbleStylList = warpObj['tableStyles']['a:tblStyleLst']['a:tblStyle']
    if (tbleStylList) {
      if (tbleStylList.constructor === Array) {
        for (let k = 0; k < tbleStylList.length; k++) {
          if (tbleStylList[k]['attrs']['styleId'] === tbleStyleId) {
            thisTblStyle = tbleStylList[k]
          }
        }
      }
      else {
        if (tbleStylList['attrs']['styleId'] === tbleStyleId) {
          thisTblStyle = tbleStylList
        }
      }
    }
  }

  let themeColor = ''
  let tbl_bgFillschemeClr = getTextByPathList(thisTblStyle, ['a:tblBg', 'a:fillRef'])
  if (tbl_bgFillschemeClr) {
    themeColor = getSolidFill(tbl_bgFillschemeClr, undefined, undefined, warpObj)
  }
  if (tbl_bgFillschemeClr === undefined) {
    tbl_bgFillschemeClr = getTextByPathList(thisTblStyle, ['a:wholeTbl', 'a:tcStyle', 'a:fill', 'a:solidFill'])
    themeColor = getSolidFill(tbl_bgFillschemeClr, undefined, undefined, warpObj)
  }
  if (themeColor !== '') themeColor = '#' + themeColor

  const trNodes = tableNode['a:tr']

  const data = []
  if (trNodes.constructor === Array) {
    for (const trNode of trNodes) {
      const tcNodes = trNode['a:tc']
      const tr = []

      if (tcNodes.constructor === Array) {
        for (const tcNode of tcNodes) {
          const text = genTextBody(tcNode['a:txBody'], tcNode, undefined, undefined, warpObj)
          const rowSpan = getTextByPathList(tcNode, ['attrs', 'rowSpan'])
          const colSpan = getTextByPathList(tcNode, ['attrs', 'gridSpan'])
          const vMerge = getTextByPathList(tcNode, ['attrs', 'vMerge'])
          const hMerge = getTextByPathList(tcNode, ['attrs', 'hMerge'])

          tr.push({ text, rowSpan, colSpan, vMerge, hMerge })
        }
      }
      else {
        const text = genTextBody(tcNodes['a:txBody'], tcNodes, undefined, undefined, warpObj)
        tr.push({ text })
      }
      data.push(tr)
    }
  }
  else {
    const tcNodes = trNodes['a:tc']
    const tr = []

    if (tcNodes.constructor === Array) {
      for (const tcNode of tcNodes) {
        const text = genTextBody(tcNode['a:txBody'], tcNode, undefined, undefined, warpObj)
        tr.push({ text })
      }
    }
    else {
      const text = genTextBody(tcNodes['a:txBody'], tcNodes, undefined, undefined, warpObj)
      tr.push({ text })
    }
    data.push(tr)
  }

  return {
    type: 'table',
    top,
    left,
    width,
    height,
    data,
    themeColor,
  }
}

export async function genChart(node, warpObj) {
  const xfrmNode = getTextByPathList(node, ['p:xfrm'])
  const { top, left } = getPosition(xfrmNode, undefined, undefined, warpObj.options.slideFactor)
  const { width, height } = getSize(xfrmNode, undefined, undefined, warpObj.options.slideFactor)

  const rid = node['a:graphic']['a:graphicData']['c:chart']['attrs']['r:id']
  const refName = warpObj['slideResObj'][rid]['target']
  const content = await readXmlFile(warpObj['zip'], refName)
  const plotArea = getTextByPathList(content, ['c:chartSpace', 'c:chart', 'c:plotArea'])

  const chart = getChartInfo(plotArea)

  if (!chart) return {}

  const data = {
    type: 'chart',
    top,
    left,
    width,
    height,
    data: chart.data,
    chartType: chart.type,
  }
  if (chart.marker !== undefined) data.marker = chart.marker
  if (chart.barDir !== undefined) data.barDir = chart.barDir
  if (chart.holeSize !== undefined) data.holeSize = chart.holeSize
  if (chart.grouping !== undefined) data.grouping = chart.grouping
  if (chart.style !== undefined) data.style = chart.style

  return data
}

export function genDiagram(node, warpObj) {
  const xfrmNode = getTextByPathList(node, ['p:xfrm'])
  const { left, top } = getPosition(xfrmNode, undefined, undefined, warpObj.options.slideFactor)
  const { width, height } = getSize(xfrmNode, undefined, undefined, warpObj.options.slideFactor)

  const dgmDrwSpArray = getTextByPathList(warpObj['digramFileContent'], ['p:drawing', 'p:spTree', 'p:sp'])
  const elements = []
  if (dgmDrwSpArray) {
    for (const item of dgmDrwSpArray) {
      const el = processSpNode(item, warpObj, 'diagramBg')
      if (el) elements.push(el)
    }
  }

  return {
    type: 'diagram',
    left,
    top,
    width,
    height,
    elements,
  }
}
