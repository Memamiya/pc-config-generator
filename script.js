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
        const fileName = file.name.toLowerCase();
        
        if (!fileName.endsWith('.xlsx') && !fileName.endsWith('.xls') && !fileName.endsWith('.csv')) {
            showStatus('Excel (.xlsx, .xls) またはCSVファイルを選択してください。', 'error');
            return;
        }

        showStatus('処理中...', 'processing');

        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                
                // 最初のシートを取得
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
                
                console.log('Data loaded:', jsonData);
                
                // データ形式を判定（縦型か横型か）
                if (isVerticalFormat(jsonData)) {
                    processVerticalData(jsonData);
                } else {
                    processHorizontalData(jsonData);
                }
            } catch (error) {
                console.error('Error:', error);
                showStatus('ファイルの読み込みに失敗しました: ' + error.message, 'error');
            }
        };
        reader.readAsArrayBuffer(file);
    }

    function isVerticalFormat(data) {
        // 最初の列に「受付番号」「モデル名」などの項目名が含まれているかチェック
        if (data.length > 0 && data[0].length >= 2) {
            const firstColumn = data.map(row => row[0] ? row[0].toString() : '');
            const verticalKeywords = ['受付番号', 'モデル名', 'プロセッサー', 'OS', 'メモリ'];
            return verticalKeywords.some(keyword => 
                firstColumn.some(cell => cell.includes(keyword))
            );
        }
        return false;
    }

    function processVerticalData(data) {
        try {
            const pcData = {};
            
            // 縦型データから情報を抽出
            for (let i = 0; i < data.length; i++) {
                const row = data[i];
                if (row.length < 2) continue;
                
                const itemName = (row[0] || '').toString().trim();
                const itemValue = (row[1] || '').toString().trim();
                
                // 項目名に基づいてデータを抽出
                if (itemName.includes('モデル名')) {
                    pcData['モデル名'] = itemValue;
                } else if (itemName.includes('プロセッサー') || itemName.includes('プロセッサ')) {
                    pcData['プロセッサー'] = itemValue;
                } else if (itemName === 'OS') {
                    pcData['OS'] = itemValue;
                } else if (itemName.includes('ディスプレイ')) {
                    pcData['ディスプレイ'] = itemValue;
                } else if (itemName.includes('Webカメラ') || itemName.includes('カメラ')) {
                    pcData['Webカメラ'] = itemValue;
                } else if (itemName.includes('メモリ')) {
                    pcData['メモリ'] = itemValue;
                } else if (itemName.includes('ストレージ')) {
                    pcData['ストレージ'] = itemValue;
                } else if (itemName.includes('グラフィックス')) {
                    pcData['グラフィックス'] = itemValue;
                } else if (itemName.includes('光学ドライブ')) {
                    pcData['光学ドライブ'] = itemValue;
                } else if (itemName.includes('無線LAN')) {
                    pcData['無線LAN/Bluetooth'] = itemValue;
                } else if (itemName.includes('キーボード')) {
                    pcData['キーボード'] = itemValue ? '有' : '無';
                } else if (itemName.includes('マウス')) {
                    pcData['マウス'] = itemValue ? '有' : '無';
                } else if (itemName.toLowerCase().includes('office') || itemName.includes('オフィス')) {
                    pcData['Officeソフト'] = itemValue ? '有' : '無';
                }
            }
            
            // デフォルト値の設定
            if (!pcData['Officeソフト']) {
                pcData['Officeソフト'] = '無';
            }
            if (!pcData['マウス']) {
                pcData['マウス'] = '無';
            }
            if (!pcData['キーボード'] && pcData['キーボード'] !== '無') {
                pcData['キーボード'] = '有';
            }
            
            console.log('Extracted data:', pcData);
            
            // Excelファイルを生成
            const wb = createWorkbook(pcData);
            const modelName = pcData['モデル名'] || 'PC構成表';
            const fileName = `${modelName.replace(/[<>:"/\\|?*]/g, '_')}_構成表.xlsx`;
            
            XLSX.writeFile(wb, fileName);
            showStatus('ダウンロードが完了しました！', 'success');
            
        } catch (error) {
            console.error('Process error:', error);
            showStatus('処理中にエラーが発生しました: ' + error.message, 'error');
        }
    }

    function processHorizontalData(data) {
        // 既存の横型データ処理（以前のコードと同じ）
        try {
            if (data.length < 2) {
                showStatus('データが見つかりませんでした。', 'error');
                return;
            }

            const headers = data[0].map(h => (h || '').toString().trim());
            const fieldMapping = {
                'モデル名': ['モデル名', 'モデル', '製品名', 'Model', '商品名', '型番'],
                'プロセッサー': ['プロセッサー', 'プロセッサ', 'CPU', 'Processor'],
                'OS': ['OS', 'オペレーティングシステム', 'Operating System'],
                'ディスプレイ': ['ディスプレイ', 'Display', '画面', 'ディスプレー'],
                'Webカメラ': ['Webカメラ', 'ウェブカメラ', 'Camera', 'カメラ'],
                'メモリ': ['メモリ', 'Memory', 'RAM', 'メモリー'],
                'ストレージ': ['ストレージ', 'Storage', 'HDD', 'SSD', 'ハードディスク'],
                'グラフィックス': ['グラフィックス', 'Graphics', 'GPU', 'グラフィック'],
                '光学ドライブ': ['光学ドライブ', 'Optical Drive', 'DVD', 'ドライブ'],
                '無線LAN/Bluetooth': ['無線LAN/Bluetooth', '無線LAN', 'Wireless', '無線', 'WiFi'],
                'キーボード': ['キーボード', 'Keyboard'],
                'マウス': ['マウス', 'Mouse']
            };

            const columnIndexes = {};
            for (const [field, possibleNames] of Object.entries(fieldMapping)) {
                for (let i = 0; i < headers.length; i++) {
                    const header = headers[i];
                    for (const name of possibleNames) {
                        if (header.includes(name)) {
                            columnIndexes[field] = i;
                            break;
                        }
                    }
                    if (columnIndexes[field] !== undefined) break;
                }
            }

            const workbooks = [];
            
            for (let i = 1; i < data.length; i++) {
                const row = data[i];
                if (!row || row.length === 0) continue;
                
                const pcData = {};
                
                for (const [field, index] of Object.entries(columnIndexes)) {
                    if (index !== undefined && row[index]) {
                        pcData[field] = row[index].toString().trim();
                    } else {
                        pcData[field] = '';
                    }
                }
                
                pcData['Officeソフト'] = '無';
                for (let j = 0; j < headers.length; j++) {
                    const header = headers[j].toLowerCase();
                    if ((header.includes('office') || header.includes('オフィス')) && 
                        row[j] && row[j].toString().trim() !== '') {
                        pcData['Officeソフト'] = '有';
                        break;
                    }
                }
                
                pcData['マウス'] = pcData['マウス'] ? '有' : '無';
                pcData['キーボード'] = pcData['キーボード'] || '有';
                
                if (pcData['モデル名'] || pcData['プロセッサー']) {
                    const wb = createWorkbook(pcData);
                    const modelName = pcData['モデル名'] || `PC_${i}`;
                    const fileName = `${modelName.replace(/[<>:"/\\|?*]/g, '_')}_構成表.xlsx`;
                    
                    workbooks.push({ workbook: wb, fileName: fileName });
                }
            }
            
            if (workbooks.length === 1) {
                XLSX.writeFile(workbooks[0].workbook, workbooks[0].fileName);
                showStatus('ダウンロードが完了しました！', 'success');
            } else if (workbooks.length > 1) {
                showStatus(`${workbooks.length}個のファイルをダウンロードします...`, 'processing');
                workbooks.forEach((wb, index) => {
                    setTimeout(() => {
                        XLSX.writeFile(wb.workbook, wb.fileName);
                        if (index === workbooks.length - 1) {
                            showStatus('すべてのダウンロードが完了しました！', 'success');
                        }
                    }, index * 1000);
                });
            } else {
                showStatus('有効なデータが見つかりませんでした。', 'error');
            }
        } catch (error) {
            console.error('Process error:', error);
            showStatus('処理中にエラーが発生しました: ' + error.message, 'error');
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
            { wch: 25 },
            { wch: 60 },
            { wch: 20 }
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
