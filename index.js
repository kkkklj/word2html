var mammoth = require("mammoth");
const fs = require('fs');

const fileName = 'andriod'
// const fileName = 'ios'

function tableContain(str) {
  const reg = /table[^>]*>[^<]*<\/table>/gi,
    startIndex = str.indexOf('<table>'),
    endIndex = str.indexOf('</table>'),
    prev = str.slice(0, startIndex),
    cur = str.slice(startIndex, endIndex + 8),
    next = str.slice(endIndex + 8);
  return `${prev}<div class="scroll-box">${cur}</div>${next}`
}

function transfer2(str) {
  str = `<!DOCTYPE html>
  <html lang="en">
  <head>
    <meta charset="UTF-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>隐私政策</title>
    <style>
      body,div,dl,dt,dd,ul,ol,li,h1,h2,h3,h4,h5,h6,pre,code,form,fieldset,legend,input,button,textarea,p,blockquote,th,td { margin:0; padding:0; }
      body { background:#fff; color:#555; font-size:14px; font-family: Verdana, Arial, Helvetica, sans-serif; }
      td,th,caption { font-size:14px; }
      h1, h2, h3, h4, h5, h6 {  font-size:100%; }
      address, caption, cite, code, dfn, em, strong, th, var { font-style:normal; }
      a { color:#555; text-decoration:none; }
      a:hover { text-decoration:underline; }
      img { border:none; }
      ol,ul,li { list-style:none; }
      input, textarea, select, button { font:14px Verdana,Helvetica,Arial,sans-serif; }
      table { border-collapse:collapse; }
      html {overflow-y: scroll;} 
      table td{border:1px solid #000}
      .scroll-box{width: calc(100vw - 30px);overflow-x: auto;}
      body{width: calc(100vw - 30px);overflow: hidden;padding: 0 15px;-webkit-text-size-adjust: 100%;}
      p{margin-top: 10px;}
    </style>
    <script>
      window.inApp = !!window.PCJSKit || !!window.PCJSKitObj;
      const win = window;
      const $callNative = (object) => {
        if (win.inApp) {
          console.log("执行协议:", object.action);
          if (win.pcBabbyFramework) {
            win.pcBabbyFramework.postMessage(JSON.stringify(object));
          } else {
            let ios = win.webkit,
              iosFn = ios.messageHandlers.pcBabbyFramework;
            ios && iosFn && iosFn.postMessage(object);
          }
        } else {
          console.warn && console.warn("协议 " + "需要在app中执行");
        }
      }
      $callNative({action:'changeTopBar',id: Math.random().toString(),callback:'globalCallBack',data:{hiddenTopBar:false,title:'隐私政策',type:0,hiddenBackButton:false,statusBarStyle:0}})
    </script>
  </head>
  <body>
    ${str}
    <script>
      const array = document.getElementsByTagName('table');
      for (let index = 0; index < array.length; index++) {
        const ele = array[index];
        const outEle = document.createElement('div'),
              parentNode = ele.parentNode;
        outEle.className = 'scroll-box';
        parentNode.replaceChild(outEle,ele);
        outEle.appendChild(ele)
      }
      const linkList = document.getElementsByTagName('a');
      while(linkList.length) {
        const p  = document.createElement('p');
        p.innerHTML = linkList[0].innerHTML;
        linkList[0].parentNode.replaceChild(p,linkList[0])
      }
    </script>
  </body>
  </html>`
  return str;
}

function unlink(filepath) {
  return new Promise((resolve, reject) => {
    var isexist = fs.existsSync( filepath )
    if(!isexist) {
      resolve();
    }
    fs.unlink(filepath, function (err) {
      if (err) {
        reject();
        throw err;
      }
      console.log('文件:' + filepath + '删除成功！');
      resolve();
    })
  })
}
const convertFn = async () => {
  const result = await mammoth.convertToHtml({ path: "./input/"+fileName+'.docx' });
  var html = result.value; // The generated HTML
  var messages = result.messages; // Any messages, such as warnings during conversion
  console.log('messages->', messages)
  html = transfer2(html);
  await unlink('./output/'+ fileName + '.html');
  fs.writeFileSync('./output/'+ fileName + '.html',html)
}

convertFn();