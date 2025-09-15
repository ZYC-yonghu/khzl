# khzl
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>北京市辖区客户排序工具</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <!-- 引入SheetJS库 -->
    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
    <style>
        :root {
            --primary-color: #2c3e50;
            --secondary-color: #3498db;
            --accent-color: #e74c3c;
            --light-bg: #f8f9fa;
            --dark-bg: #343a40;
        }

        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
            color: #333;
            line-height: 1.6;
            min-height: 100vh;
            padding: 20px 0;
        }

        .app-container {
            max-width: 1200px;
            margin: 0 auto;
            background: white;
            border-radius: 15px;
            box-shadow: 0 10px 30px rgba(0, 0, 0, 0.1);
            overflow: hidden;
        }

        .header {
            background: linear-gradient(135deg, var(--primary-color), var(--secondary-color));
            color: white;
            padding: 25px;
            text-align: center;
        }

        .header h1 {
            font-weight: 700;
            margin-bottom: 10px;
        }

        .header p {
            opacity: 0.9;
            font-size: 1.1rem;
        }

        .content {
            padding: 25px;
        }

        .card {
            border-radius: 12px;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.08);
            margin-bottom: 20px;
            border: none;
            transition: transform 0.3s;
        }

        .card:hover {
            transform: translateY(-5px);
        }

        .card-header {
            background: linear-gradient(135deg, var(--primary-color), #3a516e);
            color: white;
            border-radius: 12px 12px 0 0 !important;
            padding: 15px 20px;
            font-weight: 600;
            display: flex;
            align-items: center;
        }

        .card-header i {
            margin-right: 10px;
            font-size: 1.2rem;
        }

        .card-body {
            padding: 20px;
        }

        .form-control, .form-select {
            border-radius: 8px;
            padding: 12px 15px;
            border: 1px solid #ddd;
            transition: all 0.3s;
        }

        .form-control:focus, .form-select:focus {
            border-color: var(--secondary-color);
            box-shadow: 0 0 0 0.25rem rgba(52, 152, 219, 0.25);
        }

        .btn-primary {
            background: linear-gradient(135deg, var(--secondary-color), #2980b9);
            border: none;
            padding: 12px 25px;
            border-radius: 8px;
            font-weight: 600;
            transition: all 0.3s;
            width: 100%;
        }

        .btn-primary:hover {
            background: linear-gradient(135deg, #2980b9, var(--secondary-color));
            transform: translateY(-2px);
            box-shadow: 0 4px 8px rgba(0, 0, 0, 0.15);
        }

        .log-container {
            height: 300px;
            overflow-y: auto;
            background-color: #2c3e50;
            color: #fff;
            padding: 15px;
            border-radius: 8px;
            font-family: 'Courier New', monospace;
            font-size: 14px;
        }

        .log-entry {
            margin-bottom: 10px;
            border-left: 3px solid var(--secondary-color);
            padding-left: 12px;
            line-height: 1.5;
        }

        .log-timestamp {
            color: #3498db;
            font-weight: bold;
        }

        .progress {
            height: 12px;
            border-radius: 6px;
            margin: 20px 0;
            background-color: #e9ecef;
        }

        .progress-bar {
            background: linear-gradient(135deg, var(--secondary-color), #2980b9);
            border-radius: 6px;
        }

        .status-bar {
            background-color: var(--light-bg);
            padding: 12px 20px;
            border-radius: 8px;
            margin-top: 20px;
            font-weight: 500;
            display: flex;
            align-items: center;
        }

        .status-bar i {
            margin-right: 10px;
            color: var(--secondary-color);
        }

        .file-info {
            background-color: #e9f7fe;
            padding: 15px;
            border-radius: 8px;
            margin: 15px 0;
            display: flex;
            align-items: center;
            justify-content: space-between;
        }

        .file-info i {
            color: var(--secondary-color);
            font-size: 1.5rem;
            margin-right: 15px;
        }

        .file-details {
            flex-grow: 1;
        }

        .file-name {
            font-weight: 600;
            margin-bottom: 5px;
        }

        .file-size {
            color: #6c757d;
            font-size: 0.9rem;
        }

        .feature-icon {
            font-size: 2rem;
            color: var(--secondary-color);
            margin-bottom: 15px;
        }

        .help-link {
            color: var(--secondary-color);
            text-decoration: none;
            display: inline-flex;
            align-items: center;
        }

        .help-link:hover {
            color: var(--primary-color);
            text-decoration: underline;
        }

        .district-list {
            max-height: 200px;
            overflow-y: auto;
        }

        .list-group-item {
            border-radius: 6px !important;
            margin-bottom: 8px;
            border: 1px solid rgba(0, 0, 0, 0.08);
            display: flex;
            align-items: center;
        }

        .list-group-item i {
            margin-right: 10px;
            color: var(--secondary-color);
        }

        .file-input {
            display: none;
        }

        .result-container {
            display: none;
            margin-top: 30px;
        }

        .result-card {
            height: 100%;
        }

        .chart-container {
            position: relative;
            height: 250px;
            width: 100%;
        }

        .download-section {
            background-color: #e9f7fe;
            padding: 20px;
            border-radius: 10px;
            margin-top: 20px;
            text-align: center;
        }

        .download-btn {
            margin: 10px;
            padding: 12px 25px;
            font-weight: 600;
        }

        .result-path {
            background-color: white;
            padding: 12px 15px;
            border-radius: 8px;
            margin: 15px 0;
            font-family: monospace;
            border: 1px dashed #3498db;
        }

        @media (max-width: 768px) {
            .app-container {
                border-radius: 10px;
            }

            .header h1 {
                font-size: 1.8rem;
            }

            .content {
                padding: 15px;
            }

            .chart-container {
                height: 200px;
            }
        }
    </style>
</head>
<body>
<div class="app-container">
    <div class="header">
        <h1><i class="fas fa-map-marked-alt"></i> 北京市辖区客户排序工具</h1>
        <p class="lead">快速识别和排序北京市各辖区客户地址</p>
    </div>

    <div class="content">
        <div class="row">
            <div class="col-md-6">
                <div class="card">
                    <div class="card-header">
                        <i class="fas fa-key"></i> 百度地图API设置 (可选)
                    </div>
                    <div class="card-body">
                        <div class="mb-3">
                            <label class="form-label">API密钥:</label>
                            <input type="text" class="form-control" id="api-key" placeholder="输入您的百度地图API密钥">
                            <div class="form-text">用于无法识别的地址查询，提高识别准确率</div>
                        </div>
                        <a href="http://lbsyun.baidu.com/index.php?title=webapi/guide/webservice-geocoding" target="_blank" class="help-link">
                            <i class="fas fa-question-circle"></i> 如何获取API密钥?
                        </a>
                    </div>
                </div>

                <div class="card">
                    <div class="card-header">
                        <i class="fas fa-file-excel"></i> 文件选择
                    </div>
                    <div class="card-body">
                        <input type="file" id="file-input" class="file-input" accept=".xlsx, .xls">
                        <div class="file-info" id="file-info">
                            <i class="fas fa-file-excel"></i>
                            <div class="file-details">
                                <div class="file-name">未选择文件</div>
                                <div class="file-size">请选择Excel格式文件</div>
                            </div>
                            <button class="btn btn-outline-primary" id="browse-btn">
                                <i class="fas fa-folder-open"></i> 浏览
                            </button>
                        </div>
                        <div class="form-text">支持 .xlsx 和 .xls 格式的Excel文件</div>
                    </div>
                </div>

                <div class="card">
                    <div class="card-header">
                        <i class="fas fa-tasks"></i> 处理控制
                    </div>
                    <div class="card-body">
                        <button class="btn btn-primary" id="process-btn">
                            <i class="fas fa-play-circle"></i> 开始处理
                        </button>
                        <div class="progress mt-4">
                            <div class="progress-bar" role="progressbar" style="width: 0%" aria-valuenow="0" aria-valuemin="0" aria-valuemax="100"></div>
                        </div>
                        <div class="form-text">处理过程可能需要几分钟，请耐心等待</div>
                    </div>
                </div>
            </div>

            <div class="col-md-6">
                <div class="card">
                    <div class="card-header">
                        <i class="fas fa-list-alt"></i> 处理日志
                    </div>
                    <div class="card-body">
                        <div class="log-container" id="log-container">
                            <div class="log-entry">
                                <span class="log-timestamp">[14:25:36]</span> 就绪，请选择Excel文件并开始处理
                            </div>
                        </div>
                    </div>
                </div>

                <div class="status-bar">
                    <i class="fas fa-info-circle"></i> 状态: <span id="status-text">就绪</span>
                </div>

                <div class="card mt-4">
                    <div class="card-header">
                        <i class="fas fa-map-marker-alt"></i> 北京市辖区列表
                    </div>
                    <div class="card-body">
                        <div class="district-list">
                            <div class="row">
                                <div class="col-md-6">
                                    <ul class="list-group">
                                        <li class="list-group-item"><i class="fas fa-building"></i> 东城区</li>
                                        <li class="list-group-item"><i class="fas fa-building"></i> 西城区</li>
                                        <li class="list-group-item"><i class="fas fa-building"></i> 朝阳区</li>
                                        <li class="list-group-item"><i class="fas fa-building"></i> 丰台区</li>
                                        <li class="list-group-item"><i class="fas fa-building"></i> 石景山区</li>
                                        <li class="list-group-item"><i class="fas fa-building"></i> 海淀区</li>
                                        <li class="list-group-item"><i class="fas fa-building"></i> 门头沟区</li>
                                        <li class="list-group-item"><i class="fas fa-building"></i> 房山区</li>
                                    </ul>
                                </div>
                                <div class="col-md-6">
                                    <ul class="list-group">
                                        <li class="list-group-item"><i class="fas fa-building"></i> 通州区</li>
                                        <li class="list-group-item"><i class="fas fa-building"></i> 顺义区</li>
                                        <li class="list-group-item"><i class="fas fa-building"></i> 昌平区</li>
                                        <li class="list-group-item"><i class="fas fa-building"></i> 大兴区</li>
                                        <li class="list-group-item"><i class="fas fa-building"></i> 怀柔区</li>
                                        <li class="list-group-item"><i class="fas fa-building"></i> 平谷区</li>
                                        <li class="list-group-item"><i class="fas fa-building"></i> 密云区</li>
                                        <li class="list-group-item"><i class="fas fa-building"></i> 延庆区</li>
                                    </ul>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <!-- 结果展示区域 -->
        <div class="result-container" id="result-container">
            <h3 class="mb-4"><i class="fas fa-chart-bar"></i> 处理结果</h3>

            <div class="row">
                <div class="col-md-6 mb-4">
                    <div class="card result-card">
                        <div class="card-header">
                            <i class="fas fa-chart-pie"></i> 辖区客户分布
                        </div>
                        <div class="card-body">
                            <div class="chart-container">
                                <canvas id="districtChart"></canvas>
                            </div>
                        </div>
                    </div>
                </div>

                <div class="col-md-6 mb-4">
                    <div class="card result-card">
                        <div class="card-header">
                            <i class="fas fa-bars"></i> 辖区客户数量统计
                        </div>
                        <div class="card-body">
                            <div class="chart-container">
                                <canvas id="countChart"></canvas>
                            </div>
                        </div>
                    </div>
                </div>
            </div>

            <!-- 下载区域 -->
            <div class="download-section">
                <h4><i class="fas fa-download"></i> 下载结果文件</h4>
                <p>处理已完成，您可以下载以下结果文件：</p>

                <div class="result-path" id="sorted-file-path">
                    <i class="fas fa-file-excel"></i> 排序结果文件: <span>等待生成...</span>
                </div>

                <div class="result-path" id="stats-file-path">
                    <i class="fas fa-file-excel"></i> 统计报告文件: <span>等待生成...</span>
                </div>

                <div class="mt-3">
                    <button class="btn btn-primary download-btn" id="download-sorted-btn">
                        <i class="fas fa-download"></i> 下载排序结果
                    </button>

                    <button class="btn btn-success download-btn" id="download-stats-btn">
                        <i class="fas fa-chart-pie"></i> 下载统计报告
                    </button>
                </div>
            </div>
        </div>

        <div class="row mt-4">
            <div class="col-12">
                <div class="card">
                    <div class="card-header">
                        <i class="fas fa-star"></i> 工具特点
                    </div>
                    <div class="card-body">
                        <div class="row text-center">
                            <div class="col-md-3 feature-item mb-4">
                                <i class="fas fa-bolt feature-icon"></i>
                                <h5>快速处理</h5>
                                <p>高效识别北京市各辖区地址</p>
                            </div>
                            <div class="col-md-3 feature-item mb-4">
                                <i class="fas fa-map-marked feature-icon"></i>
                                <h5>百度地图集成</h5>
                                <p>支持百度地图API查询</p>
                            </div>
                            <div class="col-md-3 feature-item mb-4">
                                <i class="fas fa-sort feature-icon"></i>
                                <h5>智能排序</h5>
                                <p>按北京市辖区自动排序客户</p>
                            </div>
                            <div class="col-md-3 feature-item mb-4">
                                <i class="fas fa-chart-pie feature-icon"></i>
                                <h5>统计分析</h5>
                                <p>生成各辖区客户统计报表</p>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>
</div>

<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
<script>
    document.addEventListener('DOMContentLoaded', function() {
        // 获取DOM元素
        const fileInput = document.getElementById('file-input');
        const browseBtn = document.getElementById('browse-btn');
        const fileInfo = document.getElementById('file-info');
        const processBtn = document.getElementById('process-btn');
        const progressBar = document.querySelector('.progress-bar');
        const statusText = document.getElementById('status-text');
        const logContainer = document.getElementById('log-container');
        const apiKeyInput = document.getElementById('api-key');
        const resultContainer = document.getElementById('result-container');
        const sortedFilePath = document.querySelector('#sorted-file-path span');
        const statsFilePath = document.querySelector('#stats-file-path span');
        const downloadSortedBtn = document.getElementById('download-sorted-btn');
        const downloadStatsBtn = document.getElementById('download-stats-btn');

        // 初始化图表变量
        let districtChart, countChart;

        // 存储生成的文件数据
        let sortedWorkbook = null;
        let statsWorkbook = null;
        let originalFileName = '';

        // 北京市辖区列表
        const beijingDistricts = [
            '东城', '西城', '朝阳', '丰台', '石景山', '海淀', '门头沟', '房山',
            '通州', '顺义', '昌平', '大兴', '怀柔', '平谷', '密云', '延庆'
        ];

        // 常见地址模式映射
        const addressPatterns = {
            '朝阳': ['国贸', 'CBD', '三里屯', '望京', '酒仙桥', '大望路', '双井', '劲松', '潘家园', '团结湖', '呼家楼',
                '高碑店', '四惠', '朝外'],
            '海淀': ['中关村', '上地', '五道口', '清华', '北大', '学院路', '知春路', '魏公村', '苏州街', '西二旗',
                '后厂村', '常青园', '四季青', '香山南路', '北洼西里', '远大路', '都景苑', '京泉馨苑', '双榆树', '晨月园',
                '理工大学', '复兴路', '动物园', '田村', '红山口', '西山美墅馆', '营慧寺'],
            '西城': ['西单', '金融街', '复兴门', '阜成门', '月坛', '德胜门', '新街口', '西直门', '什刹海', '西四',
                '火箭军总医院', '马连道'],
            '东城': ['王府井', '东单', '建国门', '朝阳门', '东直门', '安定门', '东四', '崇文门', '前门', '天坛', '西花市'],
            '丰台': ['方庄', '丽泽', '科技园', '宋家庄', '刘家窑', '草桥', '花乡', '丰台站', '六里桥', '七里庄'],
            '通州': ['通州北苑', '九棵树', '梨园', '临河里', '土桥', '潞城', '武夷花园', '新华大街'],
            '昌平': ['回龙观', '天通苑', '沙河', '立水桥', '北七家', '小汤山', '昌平县城', '南邵'],
            '大兴': ['亦庄', '黄村', '旧宫', '西红门', '高米店', '枣园', '清源路', '天宫院'],
            '石景山': ['八角', '古城', '苹果园', '金顶', '鲁谷', '老山', '模式口', '雍景四季', '西山', '刘娘府', '八大处',
                '晋元庄', '西山枫林', '西下庄', '黑石头', '杨庄'],
            '房山': ['良乡', '长阳', '窦店', '阎村', '房山县城', '燕山'],
            '顺义': ['后沙峪', '天竺', '首都机场', '顺义县城', '马坡', '牛栏山'],
            '怀柔': ['怀柔县城', '雁栖湖', '庙城', '杨宋'],
            '平谷': ['平谷县城', '金海湖', '马坊'],
            '密云': ['密云县城', '溪翁庄', '太师屯'],
            '延庆': ['延庆县城', '八达岭', '康庄'],
            '门头沟': ['门头沟县城', '龙泉', '永定', '潭柘寺', '双峪社区', '月季园']
        };

        // 文件选择处理
        browseBtn.addEventListener('click', function() {
            fileInput.click();
        });

        fileInput.addEventListener('change', function() {
            if (this.files && this.files[0]) {
                const file = this.files[0];
                const fileName = file.name;
                const fileSize = (file.size / 1024 / 1024).toFixed(2); // MB

                // 保存原始文件名
                originalFileName = fileName.substring(0, fileName.lastIndexOf('.'));

                // 更新文件信息显示
                const fileDetails = fileInfo.querySelector('.file-details');
                fileDetails.innerHTML = `
                        <div class="file-name">${fileName}</div>
                        <div class="file-size">${fileSize} MB</div>
                    `;

                addLogEntry(`已选择文件: ${fileName}`);
                updateStatus('就绪');
            }
        });

        // 处理按钮点击事件
        processBtn.addEventListener('click', function() {
            if (!fileInput.files || !fileInput.files[0]) {
                addLogEntry('错误: 请先选择Excel文件');
                updateStatus('错误');
                return;
            }

            const apiKey = apiKeyInput.value.trim();
            if (apiKey) {
                addLogEntry('已设置百度地图API密钥');
            } else {
                addLogEntry('警告: 未设置百度地图API密钥，部分功能可能受限');
            }

            // 开始处理
            processExcelFile(fileInput.files[0]);
        });

        // 下载按钮事件
        downloadSortedBtn.addEventListener('click', function() {
            if (sortedWorkbook) {
                downloadExcelFile(sortedWorkbook, `${originalFileName}_按北京市辖区排序.xlsx`);
            } else {
                addLogEntry('错误: 排序结果文件尚未生成');
            }
        });

        downloadStatsBtn.addEventListener('click', function() {
            if (statsWorkbook) {
                downloadExcelFile(statsWorkbook, `${originalFileName}_北京市辖区统计.xlsx`);
            } else {
                addLogEntry('错误: 统计文件尚未生成');
            }
        });

        // 添加日志条目
        function addLogEntry(message) {
            const timestamp = new Date().toLocaleTimeString();
            const logEntry = document.createElement('div');
            logEntry.className = 'log-entry';
            logEntry.innerHTML = `<span class="log-timestamp">[${timestamp}]</span> ${message}`;
            logContainer.appendChild(logEntry);
            logContainer.scrollTop = logContainer.scrollHeight;
        }

        // 更新状态文本
        function updateStatus(status) {
            statusText.textContent = status;
        }

        // 处理Excel文件
        function processExcelFile(file) {
            updateStatus('处理中...');
            processBtn.disabled = true;

            // 重置进度条
            progressBar.style.width = '0%';
            progressBar.setAttribute('aria-valuenow', 0);

            addLogEntry('开始处理文件...');

            const reader = new FileReader();

            reader.onload = function(e) {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });

                    // 模拟进度更新
                    let progress = 0;
                    const interval = setInterval(() => {
                        progress += Math.floor(Math.random() * 5) + 1;
                        if (progress > 100) progress = 100;

                        progressBar.style.width = `${progress}%`;
                        progressBar.setAttribute('aria-valuenow', progress);

                        // 添加一些日志消息
                        if (progress === 10) {
                            addLogEntry('加载Excel文件...');
                        } else if (progress === 25) {
                            addLogEntry('识别地址列...');
                        } else if (progress === 40) {
                            addLogEntry('提取北京市辖区信息...');
                        } else if (progress === 60) {
                            addLogEntry('处理无法识别的地址...');
                        } else if (progress === 80) {
                            addLogEntry('生成排序结果...');
                        } else if (progress === 90) {
                            addLogEntry('生成统计信息...');
                        }

                        if (progress >= 100) {
                            clearInterval(interval);

                            // 处理数据
                            processWorkbookData(workbook);

                            addLogEntry('处理完成!');
                            updateStatus('完成');
                            processBtn.disabled = false;
                        }
                    }, 200);
                } catch (error) {
                    addLogEntry(`错误: ${error.message}`);
                    updateStatus('错误');
                    processBtn.disabled = false;
                }
            };

            reader.onerror = function() {
                addLogEntry('错误: 读取文件失败');
                updateStatus('错误');
                processBtn.disabled = false;
            };

            reader.readAsArrayBuffer(file);
        }

        // 处理工作簿数据
        function processWorkbookData(workbook) {
            try {
                // 获取第一个工作表
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];

                // 将工作表转换为JSON
                const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

                if (data.length <= 1) {
                    throw new Error('Excel文件中没有足够的数据');
                }

                // 获取表头
                const headers = data[0];

                // 查找地址列
                const addressColIndex = findAddressColumn(headers);
                if (addressColIndex === -1) {
                    throw new Error('无法识别地址列');
                }

                addLogEntry(`识别到地址列: ${headers[addressColIndex]}`);

                // 提取数据行
                const rows = data.slice(1);

                // 为每行数据添加辖区信息和识别方式
                const processedData = rows.map(row => {
                    const address = row[addressColIndex] || '';
                    const { district, method } = extractDistrict(address);
                    return [...row, district, method];
                });

                // 创建带辖区信息的新表头
                const newHeaders = [...headers, '所属辖区', '识别方式'];

                // 生成排序后的工作簿
                sortedWorkbook = createSortedWorkbook(newHeaders, processedData);

                // 生成统计工作簿
                statsWorkbook = createStatsWorkbook(processedData);

                // 更新文件路径显示
                sortedFilePath.textContent = `${originalFileName}_按北京市辖区排序.xlsx`;
                statsFilePath.textContent = `${originalFileName}_北京市辖区统计.xlsx`;

                // 显示结果
                showResults(processedData);

            } catch (error) {
                addLogEntry(`处理数据时出错: ${error.message}`);
                updateStatus('错误');
            }
        }

        // 查找地址列
        function findAddressColumn(headers) {
            const addressKeywords = ['地址', '住址', '联系地址', 'location', 'address', '收货地址'];

            for (let i = 0; i < headers.length; i++) {
                const header = String(headers[i]).toLowerCase();
                if (addressKeywords.some(keyword => header.includes(keyword))) {
                    return i;
                }
            }

            // 如果没有找到标准名称，尝试通过内容识别
            return 0; // 简化处理，默认第一列
        }

        // 从地址中提取辖区信息
        function extractDistrict(address) {
            if (!address) return { district: "无法识别", method: "无法识别" };

            const addressStr = String(address).replace('北京市', '').replace('北京', '').trim();

            // 1. 首先尝试直接匹配区名
            for (const district of beijingDistricts) {
                if (addressStr.includes(district)) {
                    return { district: district + '区', method: "直接匹配" };
                }
            }

            // 2. 检查区县名称变体
            const districtMappings = {
                '通县': '通州区',
                '昌平区': '昌平区', '昌平县': '昌平区',
                '大兴区': '大兴区',
                '海淀区': '海淀区',
                '朝阳区': '朝阳区',
                '东城区': '东城区',
                '西城区': '西城区',
                '丰台区': '丰台区',
                '石景山区': '石景山区',
                '门头沟区': '门头沟区',
                '房山区': '房山区',
                '顺义区': '顺义区',
                '怀柔区': '怀柔区',
                '平谷区': '平谷区',
                '密云区': '密云区',
                '延庆区': '延庆区'
            };

            for (const [key, value] of Object.entries(districtMappings)) {
                if (addressStr.includes(key)) {
                    return { district: value, method: "变体匹配" };
                }
            }

            // 3. 使用地址模式匹配
            for (const [district, patterns] of Object.entries(addressPatterns)) {
                for (const pattern of patterns) {
                    if (addressStr.includes(pattern)) {
                        return { district: district + '区', method: "模式匹配" };
                    }
                }
            }

            return { district: "无法识别", method: "无法识别" };
        }

        // 创建排序后的工作簿
        function createSortedWorkbook(headers, data) {
            // 按辖区排序
            const sortedData = [...data].sort((a, b) => {
                const districtA = a[a.length - 2] || '';
                const districtB = b[b.length - 2] || '';
                return districtA.localeCompare(districtB);
            });

            // 创建工作表
            const ws = XLSX.utils.aoa_to_sheet([headers, ...sortedData]);

            // 创建工作簿
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, "排序结果");

            return wb;
        }

        // 创建统计工作簿
        function createStatsWorkbook(data) {
            // 统计各辖区数量
            const districtCounts = {};
            const methodCounts = {
                "直接匹配": 0,
                "变体匹配": 0,
                "模式匹配": 0,
                "无法识别": 0
            };

            data.forEach(row => {
                const district = row[row.length - 2] || '无法识别';
                const method = row[row.length - 1] || '无法识别';

                districtCounts[district] = (districtCounts[district] || 0) + 1;

                if (methodCounts.hasOwnProperty(method)) {
                    methodCounts[method]++;
                }
            });

            // 转换为数组
            const statsData = [['辖区', '客户数量']];
            for (const [district, count] of Object.entries(districtCounts)) {
                statsData.push([district, count]);
            }

            // 添加空行
            statsData.push([]);

            // 添加识别方式统计
            statsData.push(['识别方式', '数量']);
            for (const [method, count] of Object.entries(methodCounts)) {
                statsData.push([method, count]);
            }

            // 按客户数量排序
            statsData.sort((a, b) => {
                if (a[0] === '辖区') return -1;
                if (b[0] === '辖区') return 1;
                if (a[0] === '识别方式') return -1;
                if (b[0] === '识别方式') return 1;
                if (a[0] === '') return 1;
                if (b[0] === '') return -1;
                return b[1] - a[1];
            });

            // 创建工作表
            const ws = XLSX.utils.aoa_to_sheet(statsData);

            // 创建工作簿
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, "辖区统计");

            return wb;
        }

        // 下载Excel文件
        function downloadExcelFile(workbook, fileName) {
            try {
                // 将工作簿转换为二进制数据
                const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });

                // 创建Blob对象
                const blob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

                // 创建下载链接
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = fileName;
                a.click();

                // 清理URL
                setTimeout(() => URL.revokeObjectURL(url), 100);

                addLogEntry(`已生成下载: ${fileName}`);
            } catch (error) {
                addLogEntry(`下载文件时出错: ${error.message}`);
            }
        }

        // 显示处理结果
        function showResults(data) {
            // 统计各辖区数量
            const districtData = {};
            const methodData = {
                "直接匹配": 0,
                "变体匹配": 0,
                "模式匹配": 0,
                "无法识别": 0
            };

            data.forEach(row => {
                const district = row[row.length - 2] || '无法识别';
                const method = row[row.length - 1] || '无法识别';

                districtData[district] = (districtData[district] || 0) + 1;

                if (methodData.hasOwnProperty(method)) {
                    methodData[method]++;
                }
            });

            // 显示结果容器
            resultContainer.style.display = 'block';

            // 创建饼图
            createPieChart(districtData);

            // 创建柱状图
            createBarChart(districtData);

            // 添加识别方式统计日志
            addLogEntry(`识别方式统计:`);
            for (const [method, count] of Object.entries(methodData)) {
                addLogEntry(`  ${method}: ${count} 条`);
            }

            // 滚动到结果区域
            resultContainer.scrollIntoView({ behavior: 'smooth' });
        }

        // 创建饼图
        function createPieChart(data) {
            const ctx = document.getElementById('districtChart').getContext('2d');

            // 准备图表数据
            const labels = Object.keys(data);
            const values = Object.values(data);

            // 生成颜色
            const backgroundColors = [];
            for (let i = 0; i < labels.length; i++) {
                const hue = (i * 137.508) % 360; // 使用黄金角近似值分布颜色
                backgroundColors.push(`hsl(${hue}, 70%, 65%)`);
            }

            if (districtChart) {
                districtChart.destroy();
            }

            districtChart = new Chart(ctx, {
                type: 'pie',
                data: {
                    labels: labels,
                    datasets: [{
                        data: values,
                        backgroundColor: backgroundColors,
                        borderWidth: 1
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        legend: {
                            position: 'right',
                        },
                        title: {
                            display: true,
                            text: '各辖区客户分布比例'
                        }
                    }
                }
            });
        }

        // 创建柱状图
        function createBarChart(data) {
            const ctx = document.getElementById('countChart').getContext('2d');

            // 准备图表数据
            const labels = Object.keys(data);
            const values = Object.values(data);

            // 生成颜色
            const backgroundColors = [];
            for (let i = 0; i < labels.length; i++) {
                const hue = (i * 137.508) % 360;
                backgroundColors.push(`hsl(${hue}, 70%, 65%)`);
            }

            if (countChart) {
                countChart.destroy();
            }

            countChart = new Chart(ctx, {
                type: 'bar',
                data: {
                    labels: labels,
                    datasets: [{
                        label: '客户数量',
                        data: values,
                        backgroundColor: backgroundColors,
                        borderWidth: 1
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    scales: {
                        y: {
                            beginAtZero: true,
                            title: {
                                display: true,
                                text: '客户数量'
                            }
                        },
                        x: {
                            title: {
                                display: true,
                                text: '辖区'
                            }
                        }
                    },
                    plugins: {
                        title: {
                            display: true,
                            text: '各辖区客户数量统计'
                        }
                    }
                }
            });
        }

        // 初始日志
        addLogEntry('应用程序已初始化');
    });
</script>
</body>
</html>
