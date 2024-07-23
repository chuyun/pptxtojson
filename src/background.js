import {getTextByPathList} from './utils.js'
import {processNodesInSlide} from './processNode.js'

export default async function getBackground(warpObj) {
  const elements = []
  const slideLayoutContent = warpObj['slideLayoutContent']
  const slideMasterContent = warpObj['slideMasterContent']

  const nodesSldLayout = getTextByPathList(slideLayoutContent, ['p:sldLayout', 'p:cSld', 'p:spTree'])
  const nodesSldMaster = getTextByPathList(slideMasterContent, ['p:sldMaster', 'p:cSld', 'p:spTree'])

  const showMasterSp = getTextByPathList(slideLayoutContent, ['p:sldLayout', 'attrs', 'showMasterSp'])

  if (nodesSldLayout) {
    for (const nodeKey in nodesSldLayout) {
      if (nodesSldLayout[nodeKey].constructor === Array) {
        for (let i = 0; i < nodesSldLayout[nodeKey].length; i++) {
          const ph_type = getTextByPathList(nodesSldLayout[nodeKey][i], ['p:nvSpPr', 'p:nvPr', 'p:ph', 'attrs', 'type'])
          if (ph_type !== 'pic') {
            const ret = await processNodesInSlide(nodeKey, nodesSldLayout[nodeKey][i], warpObj, 'slideLayoutBg')
            if (ret) elements.push(ret)
          }
        }
      }
      else {
        const ph_type = getTextByPathList(nodesSldLayout[nodeKey], ['p:nvSpPr', 'p:nvPr', 'p:ph', 'attrs', 'type'])
        if (ph_type !== 'pic') {
          const ret = await processNodesInSlide(nodeKey, nodesSldLayout[nodeKey], warpObj, 'slideLayoutBg')
          if (ret) elements.push(ret)
        }
      }
    }
  }
  if (nodesSldMaster && ( 1 || (showMasterSp === '1' || showMasterSp))) {
    for (const nodeKey in nodesSldMaster) {
      if (nodesSldMaster[nodeKey].constructor === Array) {
        for (let i = 0; i < nodesSldMaster[nodeKey].length; i++) {
          const ret = await processNodesInSlide(nodeKey, nodesSldMaster[nodeKey][i], warpObj, 'slideMasterBg')
          if (ret) elements.push(ret)
        }
      }
      else {
        const ret = await processNodesInSlide(nodeKey, nodesSldMaster[nodeKey], warpObj, 'slideMasterBg')
        if (ret) elements.push(ret)
      }
    }
  }
  return elements
}
