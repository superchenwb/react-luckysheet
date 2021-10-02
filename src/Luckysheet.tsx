import React, { useEffect } from 'react'

function createScript(src: string) {
  return new Promise((resolve, reject) => {
    const script = document.createElement('script')
    script.type = 'text/javascript'
    script.src = src
    script.onload = resolve
    script.onerror = reject
    document.head!.appendChild(script)
  })
}

function createLink(href: string) {
  return new Promise((resolve, reject) => {
    const script = document.createElement('link')
    script.rel = 'stylesheet'
    script.href = href
    script.onload = resolve
    script.onerror = reject
    document.head!.appendChild(script)
  })
}

interface IShowtoolbarConfig {
  undoRedo?: boolean
  paintFormat?: boolean
  currencyFormat?: boolean
  percentageFormat?: boolean //百分比格式
  numberDecrease?: boolean // '减少小数位数'
  numberIncrease?: boolean // '增加小数位数
  moreFormats?: boolean // '更多格式'
  font?: boolean // '字体'
  fontSize?: boolean // '字号大小'
  bold?: boolean // '粗体 (Ctrl+B)'
  italic?: boolean // '斜体 (Ctrl+I)'
  strikethrough?: boolean // '删除线 (Alt+Shift+5)'
  underline?: boolean // '下划线 (Alt+Shift+6)'
  textColor?: boolean // '文本颜色'
  fillColor?: boolean // '单元格颜色'
  border?: boolean // '边框'
  mergeCell?: boolean // '合并单元格'
  horizontalAlignMode?: boolean // '水平对齐方式'
  verticalAlignMode?: boolean // '垂直对齐方式'
  textWrapMode?: boolean // '换行方式'
  textRotateMode?: boolean // '文本旋转方式'
  image?: boolean // '插入图片'
  link?: boolean // '插入链接'
  chart?: boolean // '图表'（图标隐藏，但是如果配置了chart插件，右击仍然可以新建图表）
  postil?: boolean //'批注'
  pivotTable?: boolean //'数据透视表'
  function?: boolean // '公式'
  frozenMode?: boolean // '冻结方式'
  sortAndFilter?: boolean // '排序和筛选'
  conditionalFormat?: boolean // '条件格式'
  dataVerification?: boolean // '数据验证'
  splitColumn?: boolean // '分列'
  screenshot?: boolean // '截图'
  findAndReplace?: boolean // '查找替换'
  protection?: boolean // '工作表保护'
  print?: boolean // '打印'
}

interface IShowsheetbarConfig {
  add?: boolean //新增sheet
  menu?: boolean //sheet管理菜单
  sheet?: boolean //sheet页显示
}

interface IShowstatisticBarConfig {
  count?: boolean // 计数栏
  view?: boolean // 打印视图
  zoom?: boolean // 缩放
}

interface ICellRightClickConfig {
  copy?: boolean // 复制
  copyAs?: boolean // 复制为
  paste?: boolean // 粘贴
  insertRow?: boolean // 插入行
  insertColumn?: boolean // 插入列
  deleteRow?: boolean // 删除选中行
  deleteColumn?: boolean // 删除选中列
  deleteCell?: boolean // 删除单元格
  hideRow?: boolean // 隐藏选中行和显示选中行
  hideColumn?: boolean // 隐藏选中列和显示选中列
  rowHeight?: boolean // 行高
  columnWidth?: boolean // 列宽
  clear?: boolean // 清除内容
  matrix?: boolean // 矩阵操作选区
  sort?: boolean // 排序选区
  filter?: boolean // 筛选选区
  chart?: boolean // 图表生成
  image?: boolean // 插入图片
  link?: boolean // 插入链接
  data?: boolean // 数据验证
  cellFormat?: boolean // 设置单元格格式
}

interface IPager {
  pageIndex?: number //当前的页码
  pageSize?: number //每页显示多少行数据
  total?: number //数据总行数
  selectOption?: number[] //允许设置每页行数的选项
}

interface ICelldata {
  r: number
  c: number
  v: {
    m: any
    v: any
    ct: any
  }
}

interface ISheetData {
  name: string //工作表名称
  color: string //工作表颜色
  index: number //工作表索引
  status: number //激活状态
  order: number //工作表的下标
  hide: number //是否隐藏
  row: number //行数
  column: number //列数
  defaultRowHeight: number //自定义行高
  defaultColWidth: number //自定义列宽
  celldata: ICelldata[] //初始化使用的单元格数据
  config: {
    merge: {} //合并单元格
    rowlen: {} //表格行高
    columnlen: {} //表格列宽
    rowhidden: {} //隐藏行
    colhidden: {} //隐藏列
    borderInfo: {} //边框
    authority: {} //工作表保护
  }
  scrollLeft: number //左右滚动条位置
  scrollTop: number //上下滚动条位置
  luckysheet_select_save: any[] //选中的区域
  calcChain: any[] //公式链
  isPivotTable: boolean //是否数据透视表
  pivotTable: {} //数据透视表设置
  filter_select: {} //筛选范围
  filter: null //筛选配置
  luckysheet_alternateformat_save: [] //交替颜色
  luckysheet_alternateformat_save_modelCustom: [] //自定义交替颜色
  luckysheet_conditionformat_save: {} //条件格式
  frozen: {} //冻结行列配置
  chart: [] //图表配置
  zoomRatio: number // 缩放比例
  image: [] //图片
  showGridLines: number //是否显示网格线
  dataVerification: {} //数据验证配置
}

export interface ILuckysheetOptions {
  title?: string
  lang?: string
  gridKey?: string
  loadUrl?: string
  loadSheetUrl?: string
  allowUpdate?: boolean
  updateUrl?: string
  updateImageUrl?: string
  data?: ISheetData[]
  column?: number
  row?: number
  autoFormatw?: boolean
  accuracy?: number
  allowCopy?: boolean
  showtoolbar?: boolean
  showtoolbarConfig?: IShowtoolbarConfig
  showinfobar?: boolean
  showsheetbar?: boolean
  showsheetbarConfig?: IShowsheetbarConfig
  showstatisticBar?: boolean
  showstatisticBarConfig?: IShowstatisticBarConfig
  enableAddRow?: boolean
  enableAddBackTop?: boolean
  myFolderUrl?: string
  devicePixelRatio?: number
  showConfigWindowResize?: boolean
  forceCalculation?: boolean
  cellRightClickConfig?: ICellRightClickConfig
  rowHeaderWidth?: number
  columnHeaderHeight?: number
  sheetFormulaBar?: boolean
  defaultFontSize?: number
  limitSheetNameLength?: boolean
  defaultSheetNameMaxLength?: number
  pager?: IPager

  hook?: {
    /** 进入单元格编辑模式之前触发 */
    cellEditBefore?: (range: any[]) => void
    /** 更新这个单元格值之前触发 */
    cellUpdateBefore?: (r: number, c: number, value: any, isRefresh: boolean) => boolean
    /** 更新这个单元格后触发 */
    cellUpdated?: (r: number, c: number, oldValue: any, value: any, isRefresh: boolean) => void
    /** 单元格渲染前触发 */
    cellRenderBefore?: (cell: any, position: any, sheet: any, ctx: any) => boolean
    /** 单元格渲染结束后触发 */
    cellRenderAfter?: (cell: any, position: any, sheet: any, ctx: any) => void
    /** 所有单元格渲染之前执行的方法 */
    cellAllRenderBefore?: (data: any, sheet: any, ctx: any) => void
    /** 行标题单元格渲染前触发 */
    rowTitleCellRenderBefore?: (rowNum: number, position: any, ctx: any) => boolean
    /** 行标题单元格渲染后触发 */
    rowTitleCellRenderAfter?: (rowNum: string, position: any, ctx: any) => boolean
    /** 列标题单元格渲染前触发 */
    columnTitleCellRenderBefore?: (columnAbc: string, position: any, ctx: any) => boolean
    /** 列标题单元格渲染后触发 */
    columnTitleCellRenderAfter?: (columnAbc: string, position: any, ctx: any) => boolean

    /** 单元格点击前的事件 */
    cellMousedownBefore?: (cell: any, position: any, sheet: any, ctx: any) => boolean
    /** 单元格点击后的事件 */
    cellMousedown?: (cell: any, position: any, sheet: any, ctx: any) => void
    /** 鼠标移动事件 */
    sheetMousemove?: (cell: any, position: any, sheet: any, moveState: any, ctx: any) => void
    /** 鼠标滚动事件 */
    sheetMouseup?: (position: any) => void

    /** 单元格点击前的事件 */
    scroll?: (cell: any, position: any, sheet: any, ctx: any) => boolean
    /** 鼠标拖拽文件 */
    cellDragStop?: (cell: any, position: any, sheet: any, ctx: any, event: any) => void

    /** 框选或者设置选区后触发 */
    rangeSelect?: (sheet: any, range: any) => void
    /** 选区粘贴前 */
    rangePasteBefore?: (range: any, data: any) => void

    /** 点击分页按钮回调函数 */
    onTogglePager?: (page: IPager) => IPager
  }
}

export interface ILuckysheetProps {
  className?: string
  style?: React.CSSProperties
  options?: ILuckysheetOptions
}

export const Luckysheet = ({ options, className, style, ...props }: ILuckysheetProps) => {
  useEffect(() => {
    let luckysheet = window['luckysheet']
    if (!luckysheet) {
      const loaded = Promise.all([
        createLink('https://cdn.jsdelivr.net/npm/luckysheet@latest/dist/plugins/css/pluginsCss.css'),
        createLink('https://cdn.jsdelivr.net/npm/luckysheet@latest/dist/plugins/plugins.css'),
        createLink('https://cdn.jsdelivr.net/npm/luckysheet@latest/dist/css/luckysheet.css'),
        createLink('https://cdn.jsdelivr.net/npm/luckysheet@latest/dist/assets/iconfont/iconfont.css'),
        createScript('https://cdn.jsdelivr.net/npm/luckysheet@latest/dist/plugins/js/plugin.js'),
        createScript('https://cdn.jsdelivr.net/npm/luckysheet@latest/dist/luckysheet.umd.js'),
      ])
      loaded.then(() => {
        luckysheet = window['luckysheet']
        luckysheet?.create({
          container: 'go-luckysheet',
          ...options,
        })
      })
    } else {
      luckysheet?.create({
        container: 'go-luckysheet',
        ...options,
      })
    }
  }, [props])
  return (
    <div
      id="go-luckysheet"
      className={className ? `${className} go-luckysheet` : 'go-luckysheet'}
      style={{ height: '100%', position: 'relative', ...style }}
      {...props}
    />
  )
}
