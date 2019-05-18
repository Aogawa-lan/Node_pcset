const express = require('express')
const template = require('art-template')
const fs = require('fs')
const bodyParser = require('body-parser')
const xlsx = require('node-xlsx')
const session = require('express-session')
const cp = require('child_process')


const app = express()

app.engine('html',require('express-art-template'))

app.use(bodyParser.urlencoded({ extended: false }))

app.use(bodyParser.json())

app.use(session({
    secret: 'keyboard cat',
    resave: false,
    saveUninitialized: true,
    cookie: { secure: false }
}));

app.post('/post', (req, res) => {
    req.session.user = req.body
    res.redirect('/xx')
})


app.get('/',(req,res) =>{
    res.render('post.html')
})

app.get('/xx',(req,res) => {
    fs.readFile(`./fileIn/${ req.session.user.username }.txt`, (err,data) => {
        if(err){
            console.log('读取失败')
            res.send(`
            <script>
                alert('读取 ${ req.session.user.username }.txt 文件失败,请检查确认文件')
                setTimeout(()=>{
                    window.location.href = '/'
                },1000)
            </script>
            `)
        }else{
            let dataString = data.toString()

            let cpu = dataString.indexOf('处理器',0)
            ,zb = dataString.indexOf('主板', 0)
            ,gpu = dataString.indexOf('显卡', 0)
            ,ram = dataString.indexOf('内存', 0)
            ,md = dataString.indexOf('主硬盘', 0)
            ,dp = dataString.indexOf('显示器', 0)
            ,dpend = dataString.indexOf('声卡', 0)

            let cpuString = dataString.slice(cpu, zb).replace("处理器              ","")
                , zbString = dataString.slice(zb, gpu).replace("主板                ", "")
                , gpuString = dataString.slice(gpu, ram).replace("显卡                ", "")
                , ramString = dataString.slice(ram, md)
                , mdString = dataString.slice(md, dp)
                , dpString = dataString.slice(dp, dpend).replace("显示器              ", "")

            //主副硬盘
            let fypStar = dataString.indexOf('--------[ 硬盘 ]', 0)
            let fypEnd = dataString.indexOf('--------[ 内存 ]', 0)

            let fypData = dataString.slice(fypStar, fypEnd)

            let cp1 = fypData.indexOf('产品',0)
            let cp2 = fypData.lastIndexOf('产品',fypEnd)
            let size1 = fypData.indexOf('大小',0)
            let size2 = fypData.lastIndexOf('大小',fypEnd)
            let sizeend1 = fypData.indexOf('硬盘已使用',0)
            let sizeend2 = fypData.indexOf('转速',0)

            let cp1String = fypData.slice(cp1, size1).replace("产品                ", "")
            let cp2String = fypData.slice(cp2, size2).replace("产品                ", "")
            let size1String = fypData.slice(size1, sizeend1).replace("大小                ", "")
            let size2String = fypData.slice(size2, sizeend2).replace("大小                ", "")


            //内存
            let ncStar = dataString.indexOf('--------[ 内存 ]', 0)
            let ncEnd = dataString.indexOf('--------[ 显卡 ]', 0)

            let ncData = dataString.slice(ncStar, ncEnd)
            // var ncnumber = (ncData.split('Channel')).length-1

            // if (ncnumber == 2){
            //     let cp3 = ncData.indexOf('ChannelA-DIMM0', 0)
            //     let cp4 = ncData.lastIndexOf('ChannelB-DIMM0', ncEnd)
            // }

            
            let cp3 = ncData.indexOf('ChannelA-DIMM0', 0)
            let cp4 = ncData.lastIndexOf('ChannelB-DIMM0', ncEnd)
            

            if(cp4 == -1 ){
                var zzrq1 = ncData.indexOf('制造日期', 0)
                var cp3String = ncData.slice(cp3, zzrq1).replace("ChannelA-DIMM0      ", "")
                //无第二个内存为undefined
            }else{
                var zzrq1 = ncData.indexOf('制造日期', 0)
                var zzrq2 = ncData.lastIndexOf('制造日期', ncEnd)
                var cp3String = ncData.slice(cp3, zzrq1).replace("ChannelA-DIMM0      ", "")
                var cp4String = ncData.slice(cp4, zzrq2).replace("ChannelB-DIMM0      ", "")
            }

            res.render('./index.html',{
                cpu: cpuString,
                zb: zbString,
                gpu: gpuString,
                ram1: cp3String,
                ram2: cp4String,
                dp: dpString,
                md: cp1String,
                mdsize: size1String,
                fyp: cp2String,
                fypsize: size2String,
                state: `${req.session.user.username } .xls 已生成`
            })

            let xlsxData = [
                {
                    name: `${ req.session.user.username }`,
                    data: [
                        [
                            '处理器',
                            '主板',
                            '显卡',
                            '内存1',
                            '内存2',
                            '显示器',
                            '主硬盘',
                            '主硬盘大小',
                            '副硬盘',
                            '副硬盘大小',
                        ],
                        [
                            cpuString,
                            zbString,
                            gpuString,
                            cp3String,
                            cp4String,
                            dpString,
                            cp1String,
                            size1String,
                            cp2String,
                            size2String
                        ]
                    ]
                }
            ]

            var buffer = xlsx.build(xlsxData);
            fs.writeFile(`./fileOut/${ req.session.user.username }.xls`, buffer, function (err) {
                if (err)
                {
                    throw err;
                }else{
                    fs.readFileSync(`./fileOut/${req.session.user.username}.xls`, (err, data) => {
                        if(err){
                            console.log(`${req.session.user.username}.xls 已生成`)
                        }else{
                            console.log(`${req.session.user.username}.xls 已存在，请检查文件夹`)
                        }
                    })
                    
                    
                }

                // 读xlsx
                // var obj = xlsx.parse("./" + "resut.xls");
                // console.log(JSON.stringify(obj));
            }
            );
        }
    })
})

app.listen(3000,(err,data) => {
    if(err){
        console.log('Sever is error')
    }else{
        console.log('app is running : 3000')
        cp.exec('start http://localhost:3000/');
    }
})
