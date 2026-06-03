/**
 * 木材检尺对比系统 - 前端逻辑
 * 处理文件上传、列映射、搜索、材积计算等交互
 */
(function () {
    'use strict';

    // ========== 全局状态 ==========
    var PREFIX = window.WOOD_PREFIX || '/check';
    var currentFileName = null;
    var pendingFileKey = null;
    var selectedNoCol = null;
    var pendingHeaders = [];
    var pendingSheets = [];
    var pendingSheetName = null;

    // ========== 初始化 ==========
    $(function () {
        // 文件输入框显示文件名
        $('#upload-input').on('change', function () {
            var fileName = $(this).val().split('\\').pop();
            $(this).next('.custom-file-label').text(fileName || '选择 Excel 文件...');
        });

        // 上传表单提交
        $('#upload-form').on('submit', function (e) {
            e.preventDefault();
            uploadFile();
        });

        // 搜索按钮 + 防抖输入
        $('#search-btn').on('click', doSearch);
        var searchTimer = null;
        $('#search-input').on('input', function () {
            clearTimeout(searchTimer);
            searchTimer = setTimeout(doSearch, 300);
        });

        // 回车搜索
        $('#search-input').on('keydown', function (e) {
            if (e.key === 'Enter') { e.preventDefault(); doSearch(); }
        });

        // 计算表单提交
        $('#calc-form').on('submit', function (e) {
            e.preventDefault();
            doCalc();
        });

        // 加载按钮（事件委托）
        $('#file-list').on('click', '.btn-load', function () {
            loadFile($(this).data('filename'));
        });

        // 重命名按钮
        $('#file-list').on('click', '.btn-rename', function () {
            renameFile($(this).data('filename'));
        });

        // 删除按钮
        $('#file-list').on('click', '.btn-delete', function () {
            deleteFile($(this).data('filename'));
        });

        // 选择按钮（事件委托）
        $('#results-tbody').on('click', 'tr', function () {
            var recordId = $(this).data('id');
            if (recordId) selectRecord(recordId);
        });

        // 取消选择
        $('#detail-modal').on('hidden.bs.modal', function () {
            $('#results-tbody tr').removeClass('table-primary');
        });

        // 清除搜索结果
        $('#clear-results-btn').on('click', function () {
            $('#results-card').hide();
            $('#results-tbody').empty();
            $('#no-results-msg').hide();
            $('#search-input').val('');
            $('#search-status').empty();
        });

        // 自动检测按钮
        $('#col-auto-detect-btn').on('click', autoDetectColumns);
        // 列映射确认/取消
        $('#col-confirm-btn').on('click', confirmUpload);
        $('#col-cancel-btn').on('click', cancelColSelect);

        // 重命名提交
        $('#rename-form').on('submit', function (e) {
            e.preventDefault();
            submitRename();
        });

        // Sheet 切换
        $('#col-select-card').on('change', '#sheet-select', function () {
            var newSheet = $(this).val();
            if (newSheet === pendingSheetName) return;

            $('#col-status').removeClass().addClass('text-info').text('正在加载...');

            $.ajax({
                url: PREFIX + '/upload/preview_sheet',
                type: 'POST',
                contentType: 'application/json',
                data: JSON.stringify({ file_key: pendingFileKey, sheet_name: newSheet }),
                success: function (res) {
                    if (res.ok) {
                        pendingSheetName = newSheet;
                        pendingHeaders = res.headers;
                        rebuildColMapTable(res.headers);
                        $('#col-status').text('');
                    } else {
                        $('#col-status').removeClass().addClass('text-danger').text(res.error || '加载失败');
                        $('#sheet-select').val(pendingSheetName);
                    }
                },
                error: function () {
                    $('#col-status').removeClass().addClass('text-danger').text('加载失败');
                    $('#sheet-select').val(pendingSheetName);
                }
            });
        });

        // 列映射 select 变化时去掉蓝色标记
        $('#col-map-tbody').on('change', '.col-field-select', function () {
            $('#col-map-tbody .col-field-select').removeClass('border-info');
        });

        // 导出按钮
        $('#export-btn').on('click', function () {
            if (currentFileName) {
                window.open(PREFIX + '/export?file_name=' + encodeURIComponent(currentFileName), '_blank');
            }
        });

        // ===== 折叠面板 =====
        $('#toggle-panel-btn').on('click', function () {
            var panel = $('#data-panel');
            if (panel.is(':visible')) {
                panel.slideUp(200);
                $('#toggle-panel-icon').html('&#9654;');
                $('#toggle-panel-text').text('展开');
            } else {
                panel.slideDown(200);
                $('#toggle-panel-icon').html('&#9660;');
                $('#toggle-panel-text').text('收起');
            }
        });
    });

    // ========== 文件上传（两步） ==========
    function uploadFile() {
        var fileInput = $('#upload-input')[0];
        if (!fileInput.files || !fileInput.files[0]) {
            showUploadStatus('请先选择文件', 'text-danger');
            return;
        }

        var formData = new FormData();
        formData.append('file', fileInput.files[0]);

        $('#upload-btn').prop('disabled', true).text('上传中...');
        showUploadStatus('正在上传...', 'text-info');
        cancelColSelect();

        $.ajax({
            url: PREFIX + '/upload',
            type: 'POST',
            data: formData,
            processData: false,
            contentType: false,
            success: function (res) {
                if (res.ok && res.stage === 'preview') {
                    pendingFileKey = res.file_key;
                    showColSelect(res);
                } else {
                    showUploadStatus(res.error || '上传失败', 'text-danger');
                }
            },
            error: function (xhr) {
                var msg = '上传失败';
                try { msg = JSON.parse(xhr.responseText).error || msg; } catch (e) {}
                showUploadStatus(msg, 'text-danger');
            },
            complete: function () {
                $('#upload-btn').prop('disabled', false).text('上传并预览');
            }
        });
    }

    function showUploadStatus(msg, cls) {
        $('#upload-status').removeClass().addClass(cls).text(msg);
    }

    // ========== 列映射字段定义 ==========
    var FIELD_DEFS = [
        { field: 'no',              label: '编号',           required: true },
        { field: 'especie',         label: '材种' },
        { field: 'english_code',    label: '英文代码' },
        { field: 'diameter_1',      label: '直径1' },
        { field: 'diameter_2',      label: '直径2' },
        { field: 'diameter_3',      label: '直径3' },
        { field: 'diameter_4',      label: '直径4' },
        { field: 'diameter_avg',    label: '综合直径' },
        { field: 'length_m',        label: '长度' },
        { field: 'volume_m3',       label: '材积' },
        { field: 'customer',        label: '客户' },
        { field: 'is_transshipment',label: '是否转口' },
    ];

    // ========== 关键词自动匹配 ==========
    var FIELD_KEYWORDS = [
        { field: 'no',              kw: ['编号', '码单号', '序号', '顺序', 'no.', 'NO.', '号码'] },
        { field: 'diameter_1',      kw: ['径1', '直径1', '检尺1', '小头1', '径一', 'D1'] },
        { field: 'diameter_2',      kw: ['径2', '直径2', '检尺2', '小头2', '径二', 'D2'] },
        { field: 'diameter_3',      kw: ['径3', '直径3', '检尺3', '径三', 'D3'] },
        { field: 'diameter_4',      kw: ['径4', '直径4', '检尺4', '径四', 'D4'] },
        { field: 'diameter_avg',    kw: ['平均径', '综合径', '平均直径', '综合直径', '径级'] },
        { field: 'length_m',        kw: ['长度(m)', '长度(M)', '长度', 'length', '长(M)', '长(m)'] },
        { field: 'volume_m3',       kw: ['材积(m', '材积(M', '材积', '体积', 'm3', 'M3', '立方'] },
        { field: 'especie',         kw: ['材种', '树种', '木种', '品种', 'especie', '品名', '木材名称'] },
        { field: 'english_code',    kw: ['英文', '代码', 'code', 'english', '代号', '编码'] },
        { field: 'customer',        kw: ['客户', '买主', '买方', 'customer', '公司', '收货单位', '单位'] },
        { field: 'is_transshipment',kw: ['转口', '中转', '转运', 'trans', '是否转'] },
        // 直径兜底（最后匹配）
        { field: 'diameter_avg',    kw: ['直径', '径'] },
    ];

    function detectColumn(header) {
        if (!header) return null;
        var h = header.toLowerCase();
        // 先匹配精确关键词（排除兜底的直径/径）
        for (var i = 0; i < FIELD_KEYWORDS.length - 1; i++) {
            var kws = FIELD_KEYWORDS[i].kw;
            for (var j = 0; j < kws.length; j++) {
                if (h.indexOf(kws[j].toLowerCase()) !== -1) {
                    return FIELD_KEYWORDS[i].field;
                }
            }
        }
        // 兜底：直径
        var fallback = FIELD_KEYWORDS[FIELD_KEYWORDS.length - 1];
        for (var j = 0; j < fallback.kw.length; j++) {
            if (h.indexOf(fallback.kw[j].toLowerCase()) !== -1) {
                return fallback.field;
            }
        }
        return null;
    }

    // ========== 列映射展示 ==========
    function showColSelect(res) {
        var headers = res.headers;
        pendingHeaders = headers;
        pendingSheets = res.sheets || [];
        pendingSheetName = pendingSheets.length > 0 ? pendingSheets[0].name : null;

        $('#col-select-card').show();
        $('#col-file-label').text(res.file_key);
        $('#col-status').text('');

        // Sheet 选择区域
        var $sheetArea = $('#sheet-select-area');
        if (pendingSheets.length > 1) {
            var sheetOpts = '';
            for (var s = 0; s < pendingSheets.length; s++) {
                var selected = pendingSheets[s].name === pendingSheetName ? ' selected' : '';
                sheetOpts += '<option value="' + escapeHtml(pendingSheets[s].name) + '"' + selected + '>' +
                             escapeHtml(pendingSheets[s].name) + ' (' + pendingSheets[s].row_count + ' 行)</option>';
            }
            $sheetArea.html(
                '<div class="form-inline mb-2 px-2">' +
                '<label class="mr-2"><strong>Sheet:</strong></label>' +
                '<select class="form-control form-control-sm" id="sheet-select" style="min-width:200px;">' + sheetOpts + '</select>' +
                '</div>'
            ).show();
        } else {
            $sheetArea.hide().empty();
        }

        rebuildColMapTable(headers);
    }

    function rebuildColMapTable(headers) {
        var tbody = $('#col-map-tbody');
        tbody.empty();

        var optHtml = '<option value="">-- 忽略 --</option>';
        for (var f = 0; f < FIELD_DEFS.length; f++) {
            optHtml += '<option value="' + FIELD_DEFS[f].field + '">' +
                       FIELD_DEFS[f].label + '</option>';
        }

        var usedFields = {};

        for (var i = 0; i < headers.length; i++) {
            var header = headers[i] || ('列' + (i + 1));
            var autoField = detectColumn(header);

            if (autoField && usedFields[autoField]) {
                autoField = null;
            }

            var tr = $('<tr data-col="' + i + '">');
            tr.append('<td class="text-center text-muted small">' + (i + 1) + '</td>');
            tr.append('<td>' + escapeHtml(header) + '</td>');

            var selectHtml = '<select class="form-control form-control-sm col-field-select">' + optHtml + '</select>';
            var td = $('<td>').html(selectHtml);
            tr.append(td);

            if (autoField) {
                td.find('select').val(autoField);
                td.find('select').addClass('border-info');
                usedFields[autoField] = true;
            }

            tbody.append(tr);
        }
    }

    // ========== 自动检测 ==========
    function autoDetectColumns() {
        var usedFields = {};
        $('#col-map-tbody .col-field-select').removeClass('border-info');

        $('#col-map-tbody tr').each(function () {
            var colIdx = parseInt($(this).data('col'));
            var header = pendingHeaders[colIdx] || '';
            var autoField = detectColumn(header);

            if (autoField && usedFields[autoField]) {
                autoField = null;
            }

            var select = $(this).find('.col-field-select');
            if (autoField) {
                select.val(autoField);
                select.addClass('border-info');
                usedFields[autoField] = true;
            } else {
                select.val('');
            }
        });
    }

    function cancelColSelect() {
        $('#col-select-card').hide();
        pendingFileKey = null;
        pendingHeaders = [];
        pendingSheets = [];
        pendingSheetName = null;
        $('#col-map-tbody').empty();
        $('#sheet-select-area').hide().empty();
    }

    // ========== 确认解析 ==========
    function confirmUpload() {
        if (!pendingFileKey) return;

        // 收集列映射
        var colMapping = {};
        $('#col-map-tbody tr').each(function () {
            var colIdx = parseInt($(this).data('col'));
            var field = $(this).find('.col-field-select').val();
            if (field) {
                colMapping[field] = colIdx;
            }
        });

        if (Object.keys(colMapping).length === 0) {
            alert('请至少映射一列');
            return;
        }
        if (colMapping.no === undefined) {
            alert('请必须映射「编号」列');
            return;
        }

        $('#col-confirm-btn').prop('disabled', true).text('解析中...');
        $('#col-status').removeClass().addClass('text-info').text('正在解析...');

        $.ajax({
            url: PREFIX + '/upload/confirm',
            type: 'POST',
            contentType: 'application/json',
            timeout: 30000,
            data: JSON.stringify({ file_key: pendingFileKey, col_mapping: colMapping, sheet_name: pendingSheetName }),
            success: function (res) {
                if (res.ok) {
                    $('#col-status').removeClass().addClass('text-success').text(res.message);
                    currentFileName = res.file_name;
                    cancelColSelect();
                    showSearchCard(res.file_name);
                    refreshFileList();
                    showUploadStatus(res.message, 'text-success');
                } else {
                    $('#col-status').removeClass().addClass('text-danger').text(res.error || '解析失败');
                }
            },
            error: function (xhr) {
                var msg = '解析失败';
                if (xhr.statusText === 'timeout') {
                    msg = '请求超时，请刷新页面后重试';
                } else {
                    try { msg = JSON.parse(xhr.responseText).error || msg; } catch (e) {}
                }
                $('#col-status').removeClass().addClass('text-danger').text(msg);
            },
            complete: function () {
                $('#col-confirm-btn').prop('disabled', false).text('确认并解析数据');
            }
        });
    }

    // ========== 文件列表刷新 ==========
    function refreshFileList() {
        $.getJSON(PREFIX + '/files', function (res) {
            if (!res.ok || !res.files || res.files.length === 0) {
                $('#file-list').html('<p class="text-muted mb-0">暂无已保存文件，请上传码单。</p>');
                return;
            }
            var list = $('<div></div>');
            res.files.forEach(function (f) {
                var item = $('<div class="d-flex align-items-center py-2 border-bottom file-item"></div>');
                var info = $('<span class="flex-grow-1"></span>');
                info.append($('<strong></strong>').text(f.file_name));
                info.append($('<small class="text-muted ml-2"></small>').text(f.row_count + ' 条记录'));
                info.append($('<small class="text-muted ml-2"></small>').text(f.upload_time || ''));

                var loadBtn = $('<button class="btn btn-sm btn-outline-primary mr-1 btn-load" type="button">加载</button>');
                var renameBtn = $('<button class="btn btn-sm btn-outline-secondary mr-1 btn-rename" type="button">重命名</button>');
                var deleteBtn = $('<button class="btn btn-sm btn-outline-danger btn-delete" type="button">删除</button>');
                loadBtn.data('filename', f.file_name);
                renameBtn.data('filename', f.file_name);
                deleteBtn.data('filename', f.file_name);

                item.append(info, loadBtn, renameBtn, deleteBtn);
                list.append(item);
            });
            $('#file-list').empty().append(list.children('.file-item'));
        });
    }

    // ========== 加载文件 ==========
    function loadFile(filename) {
        currentFileName = filename;
        showSearchCard(filename);
        // 清空之前的搜索结果与详情
        $('#results-card').hide();
        $('#results-tbody').empty();
        $('#no-results-msg').hide();
        $('#detail-modal').modal('hide');
        $('#calc-result').hide();
        $('#search-input').val('');
        $('#search-status').empty();
    }

    // ========== 显示搜索区 ==========
    function showSearchCard(filename) {
        $('#search-card').show();
        $('#current-file-label').text('当前文件: ' + filename);
        $('#export-btn').show();
    }

    // ========== 搜索 ==========
    function doSearch() {
        var q = $('#search-input').val().trim();
        if (!q) {
            $('#results-card').hide();
            $('#results-tbody').empty();
            $('#search-status').empty();
            return;
        }

        $('#search-status').html('<span class="text-info">搜索中...</span>');

        var params = { q: q };
        if (currentFileName) {
            params.file_name = currentFileName;
        }

        $.getJSON(PREFIX + '/search', params, function (res) {
            if (!res.ok) {
                $('#search-status').html('<span class="text-danger">' + (res.error || '搜索失败') + '</span>');
                return;
            }

            $('#search-status').html('<span class="text-success">找到 ' + res.count + ' 条记录</span>');

            if (!res.records || res.records.length === 0) {
                $('#results-card').show();
                $('#results-tbody').empty();
                $('#no-results-msg').show();
                return;
            }

            $('#no-results-msg').hide();
            $('#results-card').show();
            var tbody = $('#results-tbody');
            tbody.empty();

            res.records.forEach(function (r) {
                var tr = $('<tr data-id="' + r.id + '" style="cursor:pointer;" title="点击查看详情"></tr>');
                tr.append('<td>' + escapeHtml(r.no) + '</td>');
                tr.append('<td>' + escapeHtml(r.especie) + '</td>');
                tr.append('<td>' + escapeHtml(r.english_code) + '</td>');
                tr.append('<td>' + (r.diameter_avg || '-') + '</td>');
                tr.append('<td>' + (r.length_m || '-') + '</td>');
                tr.append('<td>' + (r.volume_m3 != null ? r.volume_m3 : '-') + '</td>');
                tr.append('<td>' + escapeHtml(r.customer) + '</td>');
                tbody.append(tr);
            });
        }).fail(function () {
            $('#search-status').html('<span class="text-danger">搜索请求失败</span>');
        });
    }

    // ========== 选择记录 ==========
    function selectRecord(recordId) {
        $('#results-tbody tr').removeClass('table-primary');
        $('#results-tbody tr').each(function () {
            if ($(this).data('id') == recordId) $(this).addClass('table-primary');
        });

        var rowData = null;
        $('#results-tbody tr').each(function () {
            if ($(this).data('id') == recordId) {
                var cells = $(this).find('td');
                rowData = {
                    id: recordId,
                    no: cells.eq(0).text(),
                    especie: cells.eq(1).text(),
                    english_code: cells.eq(2).text(),
                    diameter_avg: cells.eq(3).text(),
                    length_m: cells.eq(4).text(),
                    volume_m3: cells.eq(5).text(),
                    customer: cells.eq(6).text()
                };
                return false;
            }
        });

        if (!rowData) return;
        showDetail(rowData);
        $('#detail-modal').modal('show');
    }

    function showDetail(data) {
        $('#calc-result').hide();
        $('#detail-calc-history').hide();
        $('#calc-record-id').val(data.id);

        // 获取完整记录以读取 extra_json，用于判断哪些字段被映射
        $.getJSON(PREFIX + '/search', { q: data.no, file_name: currentFileName }, function (res) {
            var full = null;
            if (res.ok && res.records) {
                for (var i = 0; i < res.records.length; i++) {
                    if (res.records[i].id == data.id) {
                        full = res.records[i];
                        break;
                    }
                }
            }
            // 合并行数据与完整记录
            var record = full || data;
            var mapped = getMappedFields(record);

            // 基本信息区 —— 只渲染映射到的字段
            var infoHtml = '';
            var infoFields = ['no', 'especie', 'english_code', 'customer', 'diameter_avg', 'length_m', 'volume_m3'];
            for (var i = 0; i < infoFields.length; i++) {
                var f = infoFields[i];
                if (mapped.indexOf(f) === -1) continue;
                var def = FIELD_LABELS[f];
                var val = record[f];
                var display = (val == null || val === '') ? '-' : val;
                if (def.unit && display !== '-') display += ' ' + def.unit;
                var cls = 'col-md-4';
                if (i > 2) cls += ' mt-2';
                var textClass = def.highlight ? 'text-primary' : '';
                infoHtml += '<div class="' + cls + '"><small class="text-muted">' + def.label + '</small><br><strong class="' + textClass + '">' + escapeHtml(String(display)) + '</strong></div>';
            }
            $('#detail-info').html(infoHtml || '<p class="text-muted">未映射任何字段</p>');

            // 检尺直径区 —— 只在映射了 D1~D4、是否转口中至少一个时显示
            var diaFields = ['diameter_1', 'diameter_2', 'diameter_3', 'diameter_4', 'is_transshipment'];
            var hasDia = false;
            for (var j = 0; j < diaFields.length; j++) {
                if (mapped.indexOf(diaFields[j]) !== -1) { hasDia = true; break; }
            }
            if (hasDia) {
                var dHtml = '';
                for (var k = 0; k < diaFields.length; k++) {
                    var df = diaFields[k];
                    if (mapped.indexOf(df) === -1) continue;
                    var dDef = FIELD_LABELS[df];
                    var dVal = record[df];
                    var dDisplay;
                    if (dDef.format) {
                        dDisplay = dDef.format(dVal);
                    } else {
                        dDisplay = (dVal != null && dVal !== '') ? dVal : '-';
                    }
                    dHtml += '<div class="col-md-3"><small class="text-muted">' + dDef.label + '</small><br><strong>' + escapeHtml(String(dDisplay)) + '</strong></div>';
                }
                $('#detail-diameters').html(dHtml);
                $('#detail-diameters-section').show();
            } else {
                $('#detail-diameters').empty();
                $('#detail-diameters-section').hide();
            }

            // 显示上次保存的计算结果
            showSavedCalcResult(record);

            // 计算区 placeholder
            $('#calc-length').attr('placeholder', '默认 ' + record.length_m);
            var diaPlaceholder = '默认 ' + record.diameter_avg;
            // 如果映射了 D1~D4，展示它们作为提示
            var mappedDiams = [];
            if (mapped.indexOf('diameter_1') !== -1) mappedDiams.push(record.diameter_1);
            if (mapped.indexOf('diameter_2') !== -1) mappedDiams.push(record.diameter_2);
            if (mapped.indexOf('diameter_3') !== -1) mappedDiams.push(record.diameter_3);
            if (mapped.indexOf('diameter_4') !== -1) mappedDiams.push(record.diameter_4);
            if (mappedDiams.length >= 2) {
                diaPlaceholder = '如 ' + mappedDiams.join(', ');
            }
            $('#calc-diameters').attr('placeholder', diaPlaceholder);
            $('#calc-length').val('');
            $('#calc-diameters').val('');
        });
    }

    // ========== 显示上次保存的计算结果 ==========
    function showSavedCalcResult(record) {
        var extra = {};
        try {
            if (record.extra_json) {
                extra = typeof record.extra_json === 'string' ? JSON.parse(record.extra_json) : record.extra_json;
            }
        } catch (e) {}
        var cr = extra.calc_result;
        if (!cr) {
            $('#detail-calc-history').hide();
            return;
        }
        var standardLabel = cr.standard === 'national' ? '国标' : '外标';
        var sign = cr.diff > 0 ? '+' : '';
        var rateClass = cr.rate > 0 ? 'text-danger' : (cr.rate < 0 ? 'text-success' : '');
        var badgeClass = cr.rate > 0 ? 'badge-danger' : (cr.rate < 0 ? 'badge-success' : 'badge-secondary');
        var labelText = cr.rate > 0 ? '涨尺' : (cr.rate < 0 ? '缩尺' : '持平');

        var html = '<strong>上次计算结果</strong> <span class="badge badge-info">' + standardLabel + '</span>';
        html += ' <small class="text-muted">(' + (cr.calc_time || '') + ')</small><br>';
        html += '直径: ' + cr.diameter_used + ' CM | 长度: ' + cr.length_used + ' M<br>';
        html += '原材积: ' + cr.original_volume.toFixed(4) + ' M³ → 新材积: ' + cr.new_volume.toFixed(4) + ' M³<br>';
        html += '涨尺量: <strong class="' + rateClass + '">' + sign + cr.diff.toFixed(4) + ' M³</strong>';
        html += ' | 涨尺率: <strong class="' + rateClass + '">' + sign + cr.rate + '%</strong>';
        html += ' <span class="badge badge-pill ' + badgeClass + '">' + labelText + '</span>';

        $('#detail-calc-history').removeClass('alert-info alert-success alert-danger alert-secondary')
            .addClass(cr.rate > 0 ? 'alert-danger' : (cr.rate < 0 ? 'alert-success' : 'alert-secondary'))
            .html(html).show();
    }

    // ========== 字段标签映射（与后端一致） ==========
    var FIELD_LABELS = {
        'no':              { label: '顺序编号', group: 'info' },
        'especie':         { label: '材种编码', group: 'info' },
        'english_code':    { label: '英文代码', group: 'info' },
        'customer':        { label: '客户',     group: 'info' },
        'diameter_avg':    { label: '综合直径', group: 'info', unit: 'CM' },
        'length_m':        { label: '原始长度', group: 'info', unit: 'M' },
        'volume_m3':       { label: '原始材积', group: 'info', unit: 'M³', highlight: true },
        'diameter_1':      { label: '检尺1',    group: 'diameter' },
        'diameter_2':      { label: '检尺2',    group: 'diameter' },
        'diameter_3':      { label: '检尺3',    group: 'diameter' },
        'diameter_4':      { label: '检尺4',    group: 'diameter' },
        'is_transshipment':{ label: '是否转口', group: 'diameter', format: function(v) { return v ? '是' : '否'; } }
    };

    // ========== 获取记录的映射字段（向前兼容） ==========
    function getMappedFields(record) {
        var extra = {};
        try {
            if (record.extra_json) {
                extra = typeof record.extra_json === 'string' ? JSON.parse(record.extra_json) : record.extra_json;
            }
        } catch (e) {}
        // 有 mapped_fields 则用它过滤，否则显示全部字段（向前兼容旧数据）
        return extra.mapped_fields || Object.keys(FIELD_LABELS);
    }

    // ========== 材积计算 ==========
    function doCalc() {
        var recordId = $('#calc-record-id').val();
        var standard = $('#calc-standard').val();
        var lengthVal = $('#calc-length').val();
        var diametersVal = $('#calc-diameters').val();

        if (!recordId) {
            alert('请先选择一根木材');
            return;
        }

        // 解析逗号分隔的直径（同时支持中英文逗号）
        var diameters = null;
        if (diametersVal && diametersVal.trim()) {
            var parts = diametersVal.split(/[,，]/);
            diameters = [];
            for (var i = 0; i < parts.length; i++) {
                var num = parseFloat(parts[i].trim());
                if (!isNaN(num)) {
                    diameters.push(num);
                }
            }
            if (diameters.length === 0) diameters = null;
        }

        var payload = {
            record_id: parseInt(recordId),
            file_name: currentFileName,
            standard: standard,
            length: lengthVal ? parseFloat(lengthVal) : null,
            diameters: diameters
        };

        $('#calc-btn').prop('disabled', true).text('计算中...');

        $.ajax({
            url: PREFIX + '/calc',
            type: 'POST',
            contentType: 'application/json',
            data: JSON.stringify(payload),
            success: function (res) {
                if (res.ok) {
                    showCalcResult(res);
                    // 刷新上次计算结果区
                    refreshCalcHistory();
                } else {
                    alert(res.error || '计算失败');
                }
            },
            error: function (xhr) {
                var msg = '计算失败';
                try { msg = JSON.parse(xhr.responseText).error || msg; } catch (e) {}
                alert(msg);
            },
            complete: function () {
                $('#calc-btn').prop('disabled', false).text('计算材积');
            }
        });
    }

    function showCalcResult(res) {
        var diff = res.diff;
        var rate = res.rate;

        var rateClass, borderClass, labelText, rateSign, badgeClass;
        if (rate > 0) {
            rateClass = 'text-danger';
            borderClass = 'border-danger';
            labelText = '涨尺';
            rateSign = '+';
            badgeClass = 'badge-danger';
        } else if (rate < 0) {
            rateClass = 'text-success';
            borderClass = 'border-success';
            labelText = '缩尺';
            rateSign = '';
            badgeClass = 'badge-success';
        } else {
            rateClass = 'text-secondary';
            borderClass = 'border-secondary';
            labelText = '持平';
            rateSign = '';
            badgeClass = 'badge-secondary';
        }

        var standardLabel = res.standard === 'national' ? '国标' : '外标';
        var diffSign = diff > 0 ? '+' : '';

        var html = '<div class="card ' + borderClass + '"><div class="card-body">';
        html += '<h6 class="card-title">计算结果 <span class="badge badge-info">' + standardLabel + '</span></h6>';
        html += '<p class="text-muted small mb-1">使用直径: ' + res.diameter_used + ' CM | 长度: ' + res.length_used + ' M</p>';
        html += '<div class="row text-center">';
        html += '<div class="col-3"><small class="text-muted">原材积</small><br><strong>' + res.original_volume.toFixed(4) + ' M³</strong></div>';
        html += '<div class="col-3"><small class="text-muted">新材积</small><br><strong>' + res.new_volume.toFixed(4) + ' M³</strong></div>';
        html += '<div class="col-3"><small class="text-muted">涨尺量</small><br><strong class="' + rateClass + '">' + diffSign + diff.toFixed(4) + ' M³</strong></div>';
        html += '<div class="col-3"><small class="text-muted">涨尺率</small><br><strong class="' + rateClass + '" style="font-size:1.4rem;">' + rateSign + rate + '%</strong></div>';
        html += '</div>';
        html += '<div class="mt-2 text-center"><span class="badge badge-pill ' + badgeClass + '">' + labelText + '</span></div>';
        html += '</div></div>';

        $('#calc-result').html(html).show();
    }

    // ========== 刷新历史计算结果 ==========
    function refreshCalcHistory() {
        var recordId = $('#calc-record-id').val();
        if (!recordId || !currentFileName) return;
        $.getJSON(PREFIX + '/search', { q: $('#search-input').val().trim(), file_name: currentFileName }, function (res) {
            if (res.ok && res.records) {
                for (var i = 0; i < res.records.length; i++) {
                    if (res.records[i].id == recordId) {
                        showSavedCalcResult(res.records[i]);
                        return;
                    }
                }
            }
        });
    }

    // ========== 删除文件 ==========
    function deleteFile(filename) {
        if (!confirm('确定删除文件 "' + filename + '" 及其所有数据？此操作不可恢复。')) {
            return;
        }

        $.ajax({
            url: PREFIX + '/files/delete',
            type: 'POST',
            contentType: 'application/json',
            data: JSON.stringify({ file_name: filename }),
            success: function (res) {
                if (res.ok) {
                    if (currentFileName === filename) {
                        currentFileName = null;
                        $('#search-card').hide();
                        $('#results-card').hide();
                        $('#detail-modal').modal('hide');
                        $('#calc-result').hide();
                    }
                    refreshFileList();
                } else {
                    alert(res.error || '删除失败');
                }
            },
            error: function () {
                alert('删除请求失败');
            }
        });
    }

    // ========== 重命名文件 ==========
    function renameFile(oldName) {
        $('#rename-old-name').val(oldName);
        $('#rename-new-name').val(oldName);
        $('#rename-status').removeClass().addClass('small').text('');
        $('#rename-submit-btn').prop('disabled', false).text('保存');
        $('#rename-modal').modal('show');
        setTimeout(function () {
            $('#rename-new-name').trigger('focus').trigger('select');
        }, 300);
    }

    function submitRename() {
        var oldName = $('#rename-old-name').val();
        var newName = $('#rename-new-name').val();
        if (!newName || newName.trim() === '') return;
        newName = newName.trim();
        if (newName === oldName) {
            $('#rename-modal').modal('hide');
            return;
        }

        $('#rename-submit-btn').prop('disabled', true).text('保存中...');
        $('#rename-status').removeClass().addClass('small text-info').text('正在重命名...');
        $.ajax({
            url: PREFIX + '/files/rename',
            type: 'POST',
            contentType: 'application/json',
            data: JSON.stringify({ old_name: oldName, new_name: newName }),
            success: function (res) {
                if (res.ok) {
                    if (currentFileName === oldName) {
                        currentFileName = newName;
                        $('#current-file-label').text('当前文件: ' + newName);
                    }
                    refreshFileList();
                    $('#rename-modal').modal('hide');
                } else {
                    $('#rename-status').removeClass().addClass('small text-danger').text(res.error || '重命名失败');
                }
            },
            error: function (xhr) {
                var msg = '重命名请求失败';
                try { msg = JSON.parse(xhr.responseText).error || msg; } catch (e) {}
                $('#rename-status').removeClass().addClass('small text-danger').text(msg);
            },
            complete: function () {
                $('#rename-submit-btn').prop('disabled', false).text('保存');
            }
        });
    }

    // ========== 工具函数 ==========
    function escapeHtml(str) {
        if (!str) return '';
        var div = document.createElement('div');
        div.appendChild(document.createTextNode(str));
        return div.innerHTML;
    }
})();
