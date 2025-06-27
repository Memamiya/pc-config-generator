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
            
            // Excelファイルを生成（デザイン適用版）
            const wb = createStyledWorkbook(pcData);
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
        // 既存の横型データ処理（省略）
        // ... 同じ処理 ...
    }

    function createStyledWorkbook(data) {
        const wb = XLSX.utils.book_new();
        
        // PC画像のプレースホルダー行を追加
        const wsData = [
            [''], // 画像用スペース（行1）
            [''], // 画像用スペース（行2）
            [''], // 画像用スペース（行3）
            [''], // 画像用スペース（行4）
            [''], // 画像用スペース（行5）
            [''], // 画像用スペース（行6）
            [''], // 画像用スペース（行7）
            [''], // 画像用スペース（行8）
            ['PC構成表'], // タイトル（行9）
            [`作成日: ${new Date().toLocaleDateString('ja-JP')}`], // 日付（行10）
            [''], // 空行（行11）
            ['項目', '内容', '備考'], // ヘッダー（行12）
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
        
        // スタイル設定（SheetJSの基本機能で可能な範囲）
        // 列幅の設定
        ws['!cols'] = [
            { wch: 25 },  // 項目列
            { wch: 60 },  // 内容列
            { wch: 20 }   // 備考列
        ];
        
        // 行の高さ設定
        ws['!rows'] = [];
        // 画像用の行を高く設定
        for (let i = 0; i < 8; i++) {
            ws['!rows'][i] = { hpt: 20 };
        }
        // タイトル行
        ws['!rows'][8] = { hpt: 30 };
        // 通常の行
        for (let i = 9; i < 30; i++) {
            ws['!rows'][i] = { hpt: 18 };
        }
        
        // セルの結合
        ws['!merges'] = [
            { s: { r: 0, c: 0 }, e: { r: 7, c: 2 } },  // 画像エリア（A1:C8）
            { s: { r: 8, c: 0 }, e: { r: 8, c: 2 } },  // タイトル（A9:C9）
            { s: { r: 9, c: 0 }, e: { r: 9, c: 2 } }   // 日付（A10:C10）
        ];
        
        // 画像プレースホルダーのテキスト
        if (ws['A1']) {
            ws['A1'].v = `[PC画像]\n${data['モデル名'] || 'PC'}の画像をここに挿入`;
        } else {
            ws['A1'] = { v: `[PC画像]\n${data['モデル名'] || 'PC'}の画像をここに挿入` };
        }
        
        // フォント設定のヒント（実際のフォント変更はExcelJSなどの高度なライブラリが必要）
        // SheetJSの基本機能では、セルの値に装飾情報を含めることは限定的
        
        // スタイル情報を含むコメント
        ws['A1'].c = [{
            a: "PC構成表ジェネレーター",
            t: "画像挿入エリア：HPの製品ページから画像をダウンロードして挿入してください。\nフォント：Meiryo UIに変更してください。"
        }];
        
        XLSX.utils.book_append_sheet(wb, ws, 'PC構成表');
        
        return wb;
    }

    // HP製品画像の検索関数（参考用）
    function getHPProductImageURL(modelName) {
        // 実際の実装では、HP APIやWebスクレイピングが必要
        // ここでは製品名からGoogle画像検索URLを生成する例
        const searchQuery = encodeURIComponent(`HP ${modelName} site:hp.com`);
        const searchURL = `https://www.google.com/search?q=${searchQuery}&tbm=isch`;
        
        // 実際の画像URLを取得するには、別途APIやサーバーサイド処理が必要
        console.log('画像検索URL:', searchURL);
        
        return null; // 現時点では画像URLを返せない
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
