import {getTextByPathList} from './utils.js'

export function extractAnimations(slideContent) {
  console.log('extractAnimations', slideContent)
  const animation = []
  const timing = getTextByPathList(slideContent, ['p:sld', 'p:timing'])
  console.log('timing', timing)
  // 时间节点列表
  const timeNodeList = getTextByPathList(timing, ['p:tnLst'])
  // 构建列表
  const buildList = getTextByPathList(timing, ['p:bldLst'])
  // <p:par>：并行时间节点（Parallel Time Node），表示其中的动画可以同时进行。
  // 第一级里面默认只有一项
  // 通用时间节点
  const commonTimeNode = getTextByPathList(timeNodeList, ['p:par', 'p:cTn'])

  const nodeType = getTextByPathList(commonTimeNode, ['attrs', 'nodeType'])
  if (nodeType === 'tmRoot') {
    animation.push({
      ...commonTimeNode['attrs'],
      children: [],
    })
  }

  // 理论上只有一项，而且一定是 SequenceTimeNode 类型
  const sequenceTimeNode = getTextByPathList(commonTimeNode, ['p:childTnLst', 'p:seq'])

  // 主序列节点
  const mainSequenceTimeNode = getTextByPathList(sequenceTimeNode, ['p:cTn'])

  const mainSequenceTimeNodeChild = getTextByPathList(mainSequenceTimeNode, ['p:childTnLst', 'p:par']) || []
  console.log('mainSequenceTimeNodeChild', mainSequenceTimeNodeChild)
  if (Array.isArray(mainSequenceTimeNodeChild)) {
    mainSequenceTimeNodeChild?.forEach((item) => {
      // const lastCTn = getTextByPathList(item, ['p:cTn', 'p:childTnLst', 'p:par', 'p:cTn', 'p:childTnLst', 'p:par', 'p:cTn'])
      const lastCTn = getLastCTn(getTextByPathList(item, ['p:cTn']))
      console.warn('lastCTn', lastCTn)
      animation[0].children.push({
        attrs: lastCTn['attrs'],
        animation: getTextByPathList(lastCTn, ['p:childTnLst', 'p:anim']),
        set: getTextByPathList(lastCTn, ['p:childTnLst', 'p:set']),
      })
    })
  }
  else {
    const lastCTn = getLastCTn(getTextByPathList(mainSequenceTimeNodeChild, ['p:cTn']))
    console.warn('lastCTn', lastCTn)
    animation[0].children.push({
      attrs: lastCTn['attrs'],
      animation: getTextByPathList(lastCTn, ['p:childTnLst', 'p:anim']),
      set: getTextByPathList(lastCTn, ['p:childTnLst', 'p:set']),
    })
  }
  console.log('animation', animation)
  return animation

}

function getLastCTn(obj) {
  const last = getTextByPathList(obj, ['p:childTnLst', 'p:par', 'p:cTn'])
  if (last) {
    return getLastCTn(last)
  }
  return obj

}

export function extractAnimationsFromTnLst() {

}

export function extractAnimationsFromChildTnLst() {

}
