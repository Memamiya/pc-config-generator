document.addEventListener('DOMContentLoaded', function() {
    const uploadArea = document.getElementById('uploadArea');
    const fileInput = document.getElementById('fileInput');
    const statusDiv = document.getElementById('status');

    // ドラッグ&ドロップの処理
    uploadArea.addEventListener('dragover', function(e) {
        e.preventDefault();
        uploadArea.classList.add('dragover');
    });

    uploadArea.addEventListener('dragleave', function(e) {
        e.preventDefault();
        uploadArea.classList.remove('dragover');
    });

    uploadArea.addEventListener('drop', function(e) {
        e.preventDefault();
        uploadArea.classList.remove('dragover');
        
        const files = e.dataTransfer.files;
        if (files.length > 0) {
            handleFile(files[0]);
        }
    });

    // ファイル選択の処理
    fileInput.addEventListener('change', function(e) {
        if (e.target.files.length > 0) {
            handleFile(e.target.files[0]);
        }
    });

    function handleFile(file) {
        if (!file.name.endsWith('.csv')) {
            showStatus('CSVファイルを選択してください。', 'error');
            return;
        }

        showStatus('処理中...', 'processing');

        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const csvData = e.target.result;
                processCSV(csvData);
            } catch (error) {
                showStatus('エラーが発生しました: ' + error.message, 'error');
            }
        };
        reader.readAsText(file, 'UTF-8');
    }

    function processCSV(csvData) {
        // CSVをパース
        const lines = csvData.split('\n');
        const headers = lines[0].split(',').map(h => h.trim());
        
        // 項目名のマッピング
        const fieldMapping = {
            'モデル名': ['モデル名', 'モデル', '製品名'],
            'プロセッサー': ['プロセッサー', 'プロセッサ', 'CPU'],
            'OS': ['OS', 'オペレーティングシステム'],
            'ディスプレイ': ['ディスプレイ', 'Display'],
            'Webカメラ': ['Webカメラ', 'ウェブカメラ'],
            'メモリ': ['メモリ', 'Memory'],
            'ストレージ': ['ストレージ', 'Storage'],
            'グラフィックス': ['グラフィックス', 'Graphics'],
            '光学ドライブ': ['光学ドライブ', 'Optical Drive'],
            '無線LAN/Bluetooth': ['無線LAN/Bluetooth', '無線LAN'],
            'キーボード': ['キーボード', 'Keyboard'],
            'マウス': ['マウス', 'Mouse']
        };

        // ヘッダーのインデックスを特定
        const columnIndexes = {};
        for (const [field, possibleNames] of Object.entries(fieldMapping)) {
            for (let i = 0; i < headers.length; i++) {
                for (const name of possibleNames) {
                    if (headers[i].includes(name)) {
                        columnIndexes[field] = i;
                        break;
                    }
                }
            }
        }

        // 各PCのデータを処理
        const workbooks = [];
        
        for (let i = 1; i < lines.length; i++) {
            if (lines[i].trim() === '') continue;
            
            const values = lines[i].split(',');
            const pcData = {};
            
            // データを抽出
            for (const [field, index] of Object.entries(columnIndexes)) {
                pcData[field] = values[index] ? values[index].trim() : '';
            }
            
            // Officeソフトの判定
            pcData['Officeソフト'] = '無';
            for (let j = 0; j < headers.length; j++) {
                if (headers[j].includes('Office') && values[j] && values[j].trim() !== '') {
                    pcData['Officeソフト'] = '有';
                    break;
                }
            }
            
            // マウスの判定
            if (!pcData['マウス'] || pcData['マウス'] === '') {
                pcData['マウス'] = '無';
            } else {
                pcData['マウス'] = '有';
            }
            
            // キーボードの判定
            if (!pcData['キーボード'] || pcData['キーボード'] === '') {
                pcData['キーボード'] = '無';
            } else {
                pcData['キーボード'] = '有';
            }
            
            // Excelワークブックを作成
            const wb = createWorkbook(pcData);
            const modelName = pcData['モデル名'] || `PC_${i}`;
            const fileName = `${modelName.replace(/[<>:"/\\|?*]/g, '_')}_構成表.xlsx`;
            
            workbooks.push({ workbook: wb, fileName: fileName });
        }
        
        // ファイルをダウンロード
        if (workbooks.length === 1) {
            XLSX.writeFile(workbooks[0].workbook, workbooks[0].fileName);
            showStatus('ダウンロードが完了しました！', 'success');
        } else if (workbooks.length > 1) {
            // 複数ファイルの場合は順番にダウンロード
            showStatus(`${workbooks.length}個のファイルをダウンロードします...`, 'processing');
            workbooks.forEach((wb, index) => {
                setTimeout(() => {
                    XLSX.writeFile(wb.workbook, wb.fileName);
                    if (index === workbooks.length - 1) {
                        showStatus('すべてのダウンロードが完了しました！', 'success');
                    }
                }, index * 1000); // 1秒間隔でダウンロード
            });
        } else {
            showStatus('データが見つかりませんでした。', 'error');
        }
    }

    function createWorkbook(data) {
        const wb = XLSX.utils.book_new();
        
        // データを配列形式に変換
        const wsData = [
            ['PC構成表'],
            [`作成日: ${new Date().toLocaleDateString('ja-JP')}`],
            [],
            ['項目', '内容', '備考'],
            ['モデル名', data['モデル名'] || '', ''],
            ['プロセッサー（CPU）', data['プロセッサー'] || '', ''],
            ['OS', data['OS'] || '', ''],
            ['ディスプレイ', data['ディスプレイ'] || '', ''],
            ['Officeソフト', data['Officeソフト'] || '無', ''],
            ['Webカメラ', data['Webカメラ'] || '', ''],
            ['メモリ', data['メモリ'] || '', ''],
            ['ストレージ', data['ストレージ'] || '', ''],
            ['グラフィックス', data['グラフィックス'] || '', ''],
            ['光学ドライブ', data['光学ドライブ'] || '', ''],
            ['無線LAN/Bluetooth', data['無線LAN/Bluetooth'] || '', ''],
            ['マウス', data['マウス'] || '無', ''],
            ['キーボード', data['キーボード'] || '有', '']
        ];
        
        const ws = XLSX.utils.aoa_to_sheet(wsData);
        
        // 列幅の設定
        ws['!cols'] = [
            { wch: 25 },  // 項目列
            { wch: 60 },  // 内容列
            { wch: 20 }   // 備考列
        ];
        
        XLSX.utils.book_append_sheet(wb, ws, 'PC構成表');
        
        return wb;
    }

    function showStatus(message, type) {
        statusDiv.textContent = message;
        statusDiv.className = `status ${type}`;
        statusDiv.style.display = 'block';
        
        if (type === 'success' || type === 'error') {
            setTimeout(() => {
                statusDiv.style.display = 'none';
            }, 5000);
        }
    }
});
