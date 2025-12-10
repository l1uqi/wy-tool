// 返利分析页面
export class RebatePage {
    constructor(app) {
        this.app = app;
        this.rulesMap = new Map();
        this.isRulesLoaded = false;
    }
    
    async render(container) {
        container.innerHTML = `
            <div class="page-container">
                <div class="page-header slide-up">
                    <h1 class="page-title">
                        <span class="icon">✨</span>
                        返利计算
                    </h1>
                    <p class="page-desc">
                        支持区间挂网底价的返利计算，导入挂网底价表和订单表即可快速计算返利
                    </p>
                </div>
                
                <div class="upload-section slide-up">
                    <div class="upload-group">
                        <label for="ruleFile">
                            <div style="margin-bottom: 5px; font-weight: bold;">第一步：导入挂网底价表</div>
                            <input type="file" id="ruleFile" accept=".xlsx,.xls" />
                            <div id="btnRuleFile" class="btn btn-primary">点击上传【挂网底价表】</div>
                        </label>
                    </div>

                    <div class="upload-group">
                        <label for="orderFile">
                            <div style="margin-bottom: 5px; font-weight: bold;">第二步：导入订单表</div>
                            <input type="file" id="orderFile" accept=".xlsx,.xls" disabled />
                            <div id="btnOrderFile" class="btn btn-primary disabled">点击上传【订单表】并计算</div>
                        </label>
                    </div>
                </div>

                <div id="status" class="status-box">请先导入【挂网底价表】...</div>
            </div>
            
            <style>
                .upload-group {
                    background: var(--bg-card);
                    padding: 20px;
                    border-radius: 8px;
                    box-shadow: 0 2px 4px rgba(0,0,0,0.05);
                    margin-bottom: 15px;
                }

                input[type="file"] {
                    display: none;
                }

                .status-box {
                    margin-top: 20px;
                    padding: 15px;
                    border-radius: 6px;
                    font-size: 14px;
                    line-height: 1.5;
                    background-color: var(--bg-secondary);
                    border: 1px solid var(--border-color);
                    min-height: 60px;
                    white-space: pre-wrap;
                }
            </style>
        `;
        
        this.bindEvents(container);
    }
    
    bindEvents(container) {
        // 解决按钮点击兼容性问题
        const bindClick = (btnId, inputId) => {
            const btn = document.getElementById(btnId);
            const input = document.getElementById(inputId);
            if(btn && input) {
                btn.addEventListener('click', (e) => {
                    e.preventDefault();
                    if (!btn.classList.contains('disabled')) {
                        input.value = '';
                        input.click();
                    }
                });
            }
        };
        
        bindClick('btnRuleFile', 'ruleFile');
        bindClick('btnOrderFile', 'orderFile');
        
        const ruleInput = document.getElementById('ruleFile');
        const orderInput = document.getElementById('orderFile');
        const btnOrder = document.getElementById('btnOrderFile');
        const status = document.getElementById('status');
        
        if (ruleInput) {
            ruleInput.addEventListener('change', (e) => this.handleRuleFile(e));
        }
        
        if (orderInput) {
            orderInput.addEventListener('change', (e) => this.handleOrderFile(e));
        }
    }
    
    /**
     * 常量定义
     */
    get CONFIG() {
        return {
            SHEET_NAMES: {
                RULES: '返利匹配规则'
            },
            COLUMNS: {
                // 订单表列名（用于逻辑判断）
                ORDER_DATE: "出库时间",
                ORDER_NO: "销售单号",
                CUST_TYPE: "客户性质",
                PROD_CODE: "商品编码",
                QTY: "销售数量",
                SALE_PRICE: "销售单价/积分",
                PAY_AMT: "支付金额",
                RECHARGE: "充值抵扣",
                REBATE_DEDUCT: "返利抵扣",
                DISCOUNT: "全部优惠金额",
                COUPON: "优惠券优惠",
                // 规则表列名
                RULE_CODE: "商品编码",
                RULE_RANGE: "区间规则",
                RULE_PRICE: "挂网底价",
                RULE_REBATE: "返利",
                RULE_MIN_QTY: "起购数量",
                RULE_CHAIN_MIN: "连锁起购量",
                RULE_SINGLE_MIN: "单体起购量",
                RULE_IS_BOMB: "是否爆品",
                RULE_GIFT_ALLOW: "赠品是否参与返利",
                RULE_START_DATE: "生效开始日期",
                RULE_END_DATE: "生效结束日期"
            },
            VALUES: {
                GIFT_PRICE: 0.01,
                DEFAULT_CHAIN_MIN: 300,
                DEFAULT_SINGLE_MIN: 120,
                BOMB_TOP_N: 3
            },
            // 最终导出时保留的列（按顺序）
            EXPORT_COLUMNS: [
                "出库时间", "签收时间", "销售单号", "外部单号", "订单类型", 
                "客户编码", "客户名称", "所属主店", "管理机构", "客户性质", 
                "商品编码", "通用名", "销售数量", "销售单价/积分", 
                "支付金额", "充值抵扣"
            ],
            // 新增的计算结果列
            OUTPUT_HEADERS: ['结算单价', '优惠券优惠', '全部优惠金额', '返利抵扣', '挂网价', '挂网底价', '单个返利金额', '是否返利', '返利金额', '不返利原因']
        };
    }
    
    /**
     * 处理规则文件上传
     */
    handleRuleFile(e) {
        const file = e.target.files[0];
        if (!file) return;

        this.updateStatus('正在解析挂网底价表...');
        
        const reader = new FileReader();
        reader.onload = evt => {
            try {
                const data = new Uint8Array(evt.target.result);
                const wb = XLSX.read(data, { type: 'array' });
                
                const sheetName = wb.SheetNames.find(name => name.trim() === this.CONFIG.SHEET_NAMES.RULES);
                if (!sheetName) {
                    throw new Error(`未找到名称为【${this.CONFIG.SHEET_NAMES.RULES}】的 Sheet`);
                }

                const ws = wb.Sheets[sheetName];
                const jsonData = XLSX.utils.sheet_to_json(ws, { header: 1 });
                
                this.parseRules(jsonData);
                
                this.isRulesLoaded = true;
                document.getElementById('orderFile').disabled = false;
                const btnOrder = document.getElementById('btnOrderFile');
                btnOrder.disabled = false;
                btnOrder.classList.remove('disabled');
                this.updateStatus(`挂网底价表解析成功！\n包含 ${this.rulesMap.size} 个商品的规则。\n请继续上传【订单表】。`);
            } catch (err) {
                console.error(err);
                this.updateStatus(`解析规则表失败: ${err.message}`, false);
                alert(`解析失败: ${err.message}`);
            }
        };
        reader.readAsArrayBuffer(file);
    }
    
    /**
     * 解析规则数据
     */
    parseRules(rows) {
        if (rows.length < 5) throw new Error("规则表格式不正确，行数过少");

        const header = rows[4];
        const C = this.CONFIG.COLUMNS;
        
        const idx = {
            code: header.indexOf(C.RULE_CODE),
            range: header.indexOf(C.RULE_RANGE),
            price: header.findIndex(h => h && h.includes('挂网') && h.includes('底价')),
            netPrice: header.findIndex(h => h && h.includes('挂网价') && !h.includes('底价')),
            rebate: header.indexOf('返利'),
            minQty: header.indexOf(C.RULE_MIN_QTY),
            chainMin: header.indexOf(C.RULE_CHAIN_MIN),
            singleMin: header.indexOf(C.RULE_SINGLE_MIN),
            isBomb: header.indexOf(C.RULE_IS_BOMB),
            giftAllow: header.indexOf(C.RULE_GIFT_ALLOW),
            start: header.indexOf(C.RULE_START_DATE),
            end: header.indexOf(C.RULE_END_DATE)
        };

        if (idx.code === -1 || idx.price === -1) {
            throw new Error("规则表缺少关键列（商品编码或挂网底价）");
        }

        this.rulesMap.clear();

        for (let i = 5; i < rows.length; i++) {
            const row = rows[i];
            const code = String(row[idx.code] || '').trim();
            const price = parseFloat(row[idx.price]);
            
            if (!code || isNaN(price)) continue;

            const range = this.parseRangeStr(row[idx.range]);
            
            const rule = {
                min: range.min,
                max: range.max,
                price: price,
                netPrice: idx.netPrice !== -1 ? parseFloat(row[idx.netPrice]) : NaN,
                rebateAmount: parseFloat(row[idx.rebate]) || 0,
                minQty: parseFloat(row[idx.minQty]),
                chainMinQty: parseFloat(row[idx.chainMin]),
                singleMinQty: parseFloat(row[idx.singleMin]),
                isBomb: (row[idx.isBomb] === '是'),
                giftAllow: (String(row[idx.giftAllow] || '').toLowerCase() === '是'),
                startDate: row[idx.start] ? new Date(row[idx.start]) : null,
                endDate: row[idx.end] ? new Date(row[idx.end]) : null
            };

            if (!this.rulesMap.has(code)) {
                this.rulesMap.set(code, []);
            }
            this.rulesMap.get(code).push(rule);
        }
    }
    
    /**
     * 处理订单文件上传
     */
    handleOrderFile(e) {
        const file = e.target.files[0];
        if (!file) return;
        if (!this.isRulesLoaded) {
            alert("请先上传挂网底价表！");
            return;
        }

        this.updateStatus('正在读取并处理订单数据...');

        const reader = new FileReader();
        reader.onload = evt => {
            try {
                const data = new Uint8Array(evt.target.result);
                const wb = XLSX.read(data, { type: 'array' });
                const ws = wb.Sheets[wb.SheetNames[0]];
                // 去掉行数限制，读取所有数据
                const json = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true });
                
                this.processOrders(json);
            } catch (err) {
                console.error(err);
                this.updateStatus(`处理订单表失败: ${err.message}`, false);
            }
        };
        reader.readAsArrayBuffer(file);
    }
    
    /**
     * 订单处理核心逻辑
     */
    processOrders(data) {
        if (!data || data.length < 2) {
            throw new Error("订单表数据为空");
        }

        const header = data[0];
        const C = this.CONFIG.COLUMNS;
        
        // 获取关键列索引（用于逻辑计算）
        const idx = {
            order: header.indexOf(C.ORDER_NO),
            code: header.indexOf(C.PROD_CODE),
            qty: header.indexOf(C.QTY),
            salePrice: header.indexOf(C.SALE_PRICE),
            pay: header.indexOf(C.PAY_AMT),
            recharge: header.indexOf(C.RECHARGE),
            rebateDeduct: header.indexOf(C.REBATE_DEDUCT),
            discount: header.indexOf(C.DISCOUNT),
            coupon: header.indexOf(C.COUPON),
            date: header.indexOf(C.ORDER_DATE),
            custType: header.indexOf(C.CUST_TYPE)
        };

        // 获取导出列的索引
        const exportIndices = this.CONFIG.EXPORT_COLUMNS.map(colName => {
            const i = header.indexOf(colName);
            return { name: colName, index: i };
        });

        // 1. 合并相同【订单号+商品编码+销售单价】的行
        const mergedRows = this.mergeOrderRows(data, idx, exportIndices);

        // 2. 预计算每个【销售单号+商品编码+销售单价】的有效总数量
        const productGroupStats = this.calculateProductGroupStats(data, idx);

        // 3. 预计算爆品统计
        const orderBombStats = this.calculateOrderBombStats(data, idx);

        // 4. 逐行计算返利（使用合并后的数据）
        const resultData = [];
        
        // 构建表头
        resultData.push([...this.CONFIG.EXPORT_COLUMNS, ...this.CONFIG.OUTPUT_HEADERS]);

        for (const [groupKey, mergedData] of mergedRows) {
            const rowData = mergedData.rowData;
            
            const groupTotalQty = productGroupStats.get(groupKey) || 0;

            // 计算返利
            const result = this.calculateRebate(rowData, groupTotalQty, orderBombStats);

            // 构建输出行：使用合并后的原始列数据
            const outputRow = [...mergedData.outputRow];

            // 返利抵扣：如果订单表中没有该列，显示为空字符串；如果有该列但值为0，显示0
            const rebateDeductValue = (mergedData.idx.rebateDeduct === -1) 
                ? '' 
                : (mergedData.totalRebateDeduct || 0);
            
            outputRow.push(
                result.settlement,
                mergedData.totalCoupon || 0,
                mergedData.totalDiscount || 0,
                rebateDeductValue,
                result.ruleNetPrice,
                result.netPrice,
                result.unitRebate,
                result.isRebate,
                result.totalRebate,
                result.reason
            );

            resultData.push(outputRow);
        }

        // 4. 导出结果
        this.downloadResult(resultData);
    }
    
    /**
     * 合并相同【订单号+商品编码+销售单价】的行
     */
    mergeOrderRows(data, idx, exportIndices) {
        const mergedMap = new Map();

        for (let i = 1; i < data.length; i++) {
            const row = data[i];
            const orderNo = row[idx.order];
            const prodCode = row[idx.code];
            const qty = parseFloat(row[idx.qty]) || 0;
            const pay = parseFloat(row[idx.pay]) || 0;
            const recharge = parseFloat(row[idx.recharge]) || 0;
            const rebateDeduct = idx.rebateDeduct !== -1 && row[idx.rebateDeduct] !== undefined && row[idx.rebateDeduct] !== null && row[idx.rebateDeduct] !== ''
                ? (parseFloat(row[idx.rebateDeduct]) || 0) 
                : 0;
            const discount = parseFloat(row[idx.discount]) || 0;
            const coupon = parseFloat(row[idx.coupon]) || 0;

            if (!orderNo || !prodCode) continue;

            const settlement = qty !== 0 ? ((pay + recharge) / qty) : 0;
            const groupKey = `${orderNo}__${prodCode}__${settlement.toFixed(2)}`;

            if (!mergedMap.has(groupKey)) {
                const outputRow = exportIndices.map(col => {
                    if (col.index === -1) return '';
                    return row[col.index];
                });

                mergedMap.set(groupKey, {
                    firstRow: row,
                    outputRow: outputRow,
                    totalQty: qty,
                    totalPay: pay,
                    totalRecharge: recharge,
                    totalRebateDeduct: rebateDeduct,
                    totalDiscount: discount,
                    totalCoupon: coupon,
                    idx: idx
                });
            } else {
                const merged = mergedMap.get(groupKey);
                merged.totalQty += qty;
                merged.totalPay += pay;
                merged.totalRecharge += recharge;
                merged.totalRebateDeduct += rebateDeduct;
                merged.totalDiscount += discount;
                merged.totalCoupon += coupon;

                const qtyColIndex = exportIndices.findIndex(col => col.name === this.CONFIG.COLUMNS.QTY);
                if (qtyColIndex !== -1) {
                    merged.outputRow[qtyColIndex] = merged.totalQty;
                }

                const payColIndex = exportIndices.findIndex(col => col.name === this.CONFIG.COLUMNS.PAY_AMT);
                if (payColIndex !== -1) {
                    merged.outputRow[payColIndex] = merged.totalPay;
                }

                const rechargeColIndex = exportIndices.findIndex(col => col.name === this.CONFIG.COLUMNS.RECHARGE);
                if (rechargeColIndex !== -1) {
                    merged.outputRow[rechargeColIndex] = merged.totalRecharge;
                }
            }
        }

        // 为每个合并后的组构建最终的 rowData
        for (const [key, merged] of mergedMap) {
            const row = merged.firstRow;
            const settlement = merged.totalQty !== 0 ? ((merged.totalPay + merged.totalRecharge) / merged.totalQty) : 0;
            
            const salePrice = (merged.idx.salePrice !== -1 && row[merged.idx.salePrice] !== undefined) 
                ? parseFloat(row[merged.idx.salePrice]) || 0 
                : 0;
            
            let orderDate = null;
            if (row[merged.idx.date]) {
                if (typeof row[merged.idx.date] === 'number') {
                    orderDate = new Date((row[merged.idx.date] - 25569) * 86400 * 1000);
                } else {
                    orderDate = new Date(row[merged.idx.date]);
                }
            }

            merged.rowData = {
                orderNo: row[merged.idx.order],
                code: String(row[merged.idx.code]),
                custType: String(row[merged.idx.custType] || ''),
                qty: Math.abs(merged.totalQty),
                rawQty: merged.totalQty,
                salePrice: salePrice || settlement,
                settlement: settlement,
                discount: merged.totalDiscount,
                recharge: merged.totalRebateDeduct,
                date: orderDate
            };
        }

        return mergedMap;
    }
    
    /**
     * 统计每个[订单+商品+销售单价]的有效总数量
     */
    calculateProductGroupStats(data, idx) {
        const stats = new Map(); 

        for (let i = 1; i < data.length; i++) {
            const row = data[i];
            const orderNo = row[idx.order];
            const prodCode = row[idx.code];
            const qty = Math.abs(parseFloat(row[idx.qty]) || 0);
            const pay = parseFloat(row[idx.pay]) || 0;
            const recharge = parseFloat(row[idx.recharge]) || 0;
            
            if (!orderNo || !prodCode) continue;

            const settlement = qty ? ((pay + recharge) / qty) : 0;
            if (Math.abs(settlement - this.CONFIG.VALUES.GIFT_PRICE) < 0.001) continue;

            const key = `${orderNo}__${prodCode}__${settlement.toFixed(2)}`;
            stats.set(key, (stats.get(key) || 0) + qty);
        }
        return stats;
    }
    
    /**
     * 统计每个订单的爆品情况
     */
    calculateOrderBombStats(rows, idx) {
        const orderStatsMap = new Map(); 

        for (let i = 1; i < rows.length; i++) {
            const row = rows[i];
            const code = String(row[idx.code]);
            const orderNo = row[idx.order];
            const qty = Math.abs(parseFloat(row[idx.qty]) || 0);
            const pay = parseFloat(row[idx.pay]) || 0;
            const recharge = parseFloat(row[idx.recharge]) || 0;
            
            const settlement = qty ? ((pay + recharge) / qty) : 0;
            if (Math.abs(settlement - this.CONFIG.VALUES.GIFT_PRICE) < 0.001) continue;

            const rules = this.rulesMap.get(code);
            if (rules && rules.length > 0 && rules[0].isBomb) {
                if (!orderStatsMap.has(orderNo)) {
                    orderStatsMap.set(orderNo, new Map());
                }
                const prodMap = orderStatsMap.get(orderNo);
                prodMap.set(code, (prodMap.get(code) || 0) + qty);
            }
        }

        const finalStats = new Map();
        for (const [orderNo, prodMap] of orderStatsMap) {
            const qtys = Array.from(prodMap.values());
            qtys.sort((a, b) => b - a); 
            const top3Sum = qtys.slice(0, this.CONFIG.VALUES.BOMB_TOP_N).reduce((sum, q) => sum + q, 0);
            finalStats.set(orderNo, top3Sum);
        }

        return finalStats;
    }
    
    /**
     * 计算单行返利结果
     */
    calculateRebate(item, groupTotalQty, orderBombStats) {
        const { code, qty, salePrice, settlement, date, custType, orderNo } = item;
        
        const rules = this.rulesMap.get(code);
        
        if (!rules || rules.length === 0) {
            return this.formatResult(item, '否', '无匹配政策', NaN, 0, false, NaN);
        }

        const firstRule = rules[0];
        
        // ✨ 优先判断1：是否是赠品
        const isGift = Math.abs(settlement - this.CONFIG.VALUES.GIFT_PRICE) < 0.001;
        if (isGift) {
            if (!firstRule.giftAllow) {
                return this.formatResult(item, '否', '赠品不参与返利', firstRule.price, firstRule.rebateAmount, false, firstRule.netPrice);
            }
            return this.formatResult(item, '是', '', firstRule.price, firstRule.rebateAmount, true, firstRule.netPrice);
        }
        
        const PRICE_TOLERANCE = 0.001;
        
        // ✨ 优先判断2：销售单价必须大于等于挂网价
        const priceToCompare = (salePrice && salePrice > 0) ? salePrice : settlement;
        if (!isNaN(firstRule.netPrice)) {
            if (priceToCompare < firstRule.netPrice - PRICE_TOLERANCE) {
                return this.formatResult(item, '否', `销售单价${priceToCompare.toFixed(2)}必须大于等于挂网价${firstRule.netPrice}（当前更低）`, firstRule.price, firstRule.rebateAmount, false, firstRule.netPrice);
            }
        }
        
        // ✨ 优先判断3：结算单价必须大于等于挂网底价
        if (settlement < firstRule.price - PRICE_TOLERANCE) {
            return this.formatResult(item, '否', `结算单价${settlement.toFixed(2)}必须大于等于底价${firstRule.price}（当前更低）`, firstRule.price, firstRule.rebateAmount, false, firstRule.netPrice);
        }
        
        // ✨ 关键逻辑：如果规则的起购数为0，价格已满足，其他条件全部跳过
        if (firstRule.minQty === 0) {
            return this.formatResult(item, '是', '', firstRule.price, firstRule.rebateAmount, false, firstRule.netPrice);
        }

        // 以下是起购数不为0时的完整判断流程
        let matchedRules = rules.filter(r => {
            if (r.startDate && r.endDate && date) {
                const d = new Date(date).setHours(0,0,0,0);
                const s = new Date(r.startDate).setHours(0,0,0,0);
                const e = new Date(r.endDate).setHours(0,0,0,0);
                return s <= d && d <= e;
            }
            return !r.startDate && !r.endDate;
        });

        const dateSpecificRules = matchedRules.filter(r => r.startDate || r.endDate);
        if (dateSpecificRules.length > 0) {
            matchedRules = dateSpecificRules.sort((a, b) => b.startDate - a.startDate);
        }

        if (matchedRules.length === 0) {
            return this.formatResult(item, '否', '无日期匹配的政策', firstRule.price, firstRule.rebateAmount, false, firstRule.netPrice);
        }

        let activeRule = matchedRules.find(r => groupTotalQty >= r.min && groupTotalQty < r.max);
        
        if (!activeRule) {
            activeRule = matchedRules[0]; 
        }

        // 计算起购量限制
        let minQtyLimit = activeRule.minQty; 
        let minQtyMsg = `起购量需≥${minQtyLimit}`;

        const isChain = custType.includes('连锁') || custType.includes('批发');
        
        if (isChain) {
            minQtyLimit = activeRule.chainMinQty || activeRule.minQty || this.CONFIG.VALUES.DEFAULT_CHAIN_MIN;
            minQtyMsg = `连锁/批发起购量需≥${minQtyLimit}`;
        } else {
            const singleLimit = activeRule.singleMinQty || activeRule.minQty;
            if (singleLimit) {
                minQtyLimit = singleLimit;
                minQtyMsg = `单体起购量需≥${singleLimit}`;
            } else {
                minQtyLimit = this.CONFIG.VALUES.DEFAULT_SINGLE_MIN;
                const bombSum = orderBombStats.get(orderNo) || 0;
                if (bombSum >= this.CONFIG.VALUES.DEFAULT_SINGLE_MIN) {
                    minQtyLimit = 0; 
                } else {
                    minQtyMsg = `单体需≥120或爆品前三合≥120(当前${bombSum})`;
                }
            }
        }

        // 起购数不为0时，执行完整的判断流程
        if (!(groupTotalQty >= activeRule.min && groupTotalQty < activeRule.max)) {
            return this.formatResult(item, '否', `总数量 ${groupTotalQty} 不在任何返利区间内`, activeRule.price, activeRule.rebateAmount, false, activeRule.netPrice);
        }

        if (activeRule.rebateAmount <= 0) {
            return this.formatResult(item, '否', '返利金额为0', activeRule.price, activeRule.rebateAmount, false, activeRule.netPrice);
        }

        if (minQtyLimit > 0 && groupTotalQty < minQtyLimit) {
            return this.formatResult(item, '否', minQtyMsg + `，当前总数${groupTotalQty}`, activeRule.price, activeRule.rebateAmount, false, activeRule.netPrice);
        }

        // 返利成功
        return this.formatResult(item, '是', '', activeRule.price, activeRule.rebateAmount, false, activeRule.netPrice);
    }
    
    formatResult(item, isRebate, reason, netPrice = NaN, unitRebate = 0, isGiftRebate = false, ruleNetPrice = NaN) {
        let totalRebate = 0;
        if (isRebate === '是') {
            if (isGiftRebate) {
                totalRebate = unitRebate * item.rawQty;
            } else {
                const discount = item.discount || 0;
                const recharge = item.recharge || 0;
                totalRebate = (unitRebate * item.rawQty) - discount - recharge;
                if (totalRebate < 0) {
                    totalRebate = 0;
                }
            }
        }

        return {
            settlement: typeof item.settlement === 'number' ? item.settlement.toFixed(2) : '',
            ruleNetPrice: isNaN(ruleNetPrice) ? '' : ruleNetPrice,
            netPrice: isNaN(netPrice) ? '' : netPrice,
            unitRebate: (typeof unitRebate === 'number') ? unitRebate.toFixed(2) : '0.00',
            isRebate,
            totalRebate: totalRebate.toFixed(2),
            reason
        };
    }
    
    /**
     * 工具函数：解析区间字符串
     */
    parseRangeStr(str) {
        if (!str) return { min: 0, max: Infinity };
        str = String(str).trim();
        if (str.includes('-')) {
            const parts = str.split('-');
            return { min: parseFloat(parts[0]), max: parseFloat(parts[1]) };
        }
        if (str.includes('+')) {
            return { min: parseFloat(str), max: Infinity };
        }
        return { min: 0, max: Infinity };
    }
    
    /**
     * 更新状态显示
     */
    updateStatus(text) {
        const status = document.getElementById('status');
        if (status) {
            status.textContent = text;
        }
    }
    
    /**
     * 下载结果
     */
    async downloadResult(data) {
        const ws = XLSX.utils.aoa_to_sheet(data);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, '返利结果');
        
        this.updateStatus('计算完成！正在准备下载...');
        
        try {
            // 检查是否在 Tauri 环境中
            if (window.__TAURI__) {
                const { invoke } = window.__TAURI__.core;
                const { save } = window.__TAURI__.dialog;
                
                // 将工作簿转换为二进制数据
                const excelBuffer = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
                
                // 打开保存文件对话框
                const filePath = await save({
                    defaultPath: `返利计算结果_${new Date().toISOString().slice(0,10).replace(/-/g, '')}_${Date.now()}.xlsx`,
                    filters: [{
                        name: 'Excel文件',
                        extensions: ['xlsx']
                    }]
                });
                
                if (filePath) {
                    // 将 ArrayBuffer 转换为 base64
                    const base64 = this.arrayBufferToBase64(excelBuffer);
                    
                    // 调用后端命令保存文件（需要修改后端支持 base64）
                    await invoke('save_excel_file', {
                        filePath: filePath,
                        content: base64
                    });
                    
                    this.updateStatus('✅ 处理完成，文件已保存！\n页面将在3秒后自动重置...');
                    
                    // 3秒后自动重置页面
                    setTimeout(() => {
                        this.resetPage();
                    }, 3000);
                } else {
                    this.updateStatus('用户取消了保存操作');
                }
            } else {
                // 非 Tauri 环境，使用浏览器下载
                XLSX.writeFile(wb, `返利计算结果_${new Date().getTime()}.xlsx`);
                this.updateStatus('✅ 处理完成，文件已下载！\n页面将在3秒后自动重置...');
                
                setTimeout(() => {
                    this.resetPage();
                }, 3000);
            }
        } catch(e) {
            console.error(e);
            this.updateStatus('❌ 下载失败: ' + e.message);
        }
    }
    
    /**
     * 将 ArrayBuffer 转换为 Base64
     */
    arrayBufferToBase64(buffer) {
        let binary = '';
        const bytes = new Uint8Array(buffer);
        const len = bytes.byteLength;
        for (let i = 0; i < len; i++) {
            binary += String.fromCharCode(bytes[i]);
        }
        return window.btoa(binary);
    }
    
    /**
     * 重置页面状态
     */
    resetPage() {
        this.rulesMap.clear();
        this.isRulesLoaded = false;
        
        document.getElementById('ruleFile').value = '';
        document.getElementById('orderFile').value = '';
        
        const orderInput = document.getElementById('orderFile');
        const btnOrder = document.getElementById('btnOrderFile');
        if (orderInput) orderInput.disabled = true;
        if (btnOrder) {
            btnOrder.disabled = true;
            btnOrder.classList.add('disabled');
        }
        
        this.updateStatus('请先导入【挂网底价表】...');
    }
}

