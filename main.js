import { utils, writeFile } from 'xlsx';

document.querySelector('#scoreForm').addEventListener('submit', (e) => {
    e.preventDefault();
    generateTable();
});

document.querySelector('#downloadBtn').addEventListener('click', () => {
    downloadExcel();
});

const judgeSelector = document.getElementById('judgeSelector');
judgeSelector.addEventListener('change', (e) => {
    const diffIndex = parseInt(e.target.value);
    if (globalJudgesData[diffIndex]) {
        renderTable(globalJudgesData[diffIndex], criteriaLabels.length);
    }
});

const criteriaBtn = document.getElementById('setCriteriaBtn');
const criteriaContainer = document.getElementById('criteriaContainer');
const groupsBtn = document.getElementById('setGroupsBtn');
const groupContainer = document.getElementById('groupContainer');

criteriaBtn.addEventListener('click', () => {
    generateCriteriaInputs();
});

groupsBtn.addEventListener('click', () => {
    generateGroupInputs();
});

let globalJudgesData = []; // Array of arrays. [Judge1Data, Judge2Data...]
let criteriaLabels = [];

function generateCriteriaInputs() {
    const count = parseInt(document.getElementById('criteriaCount').value) || 0;

    if (count <= 0) {
        criteriaContainer.innerHTML = '<div class="empty-state">올바른 심사기준 개수를 입력해주세요.</div>';
        return;
    }

    criteriaContainer.innerHTML = '';

    for (let i = 1; i <= count; i++) {
        const row = document.createElement('div');
        row.className = 'criteria-row glass-card';
        row.style.background = 'transparent';
        row.style.padding = '0';
        row.style.marginBottom = '1rem';
        row.style.borderBottom = '1px solid var(--card-border)';

        // Initial default values
        const defaultRef = 10;
        const defaultGap = 1.0;
        const hMax = defaultRef; // 10
        const hMin = defaultRef - defaultGap; // 9
        const mMax = hMin - defaultGap; // 8
        const mMin = mMax - defaultGap; // 7
        const lMax = mMin - defaultGap; // 6
        const lMin = lMax - defaultGap; // 5

        row.innerHTML = `
      <div class="criteria-label">기준 ${i}</div>
      
      <div class="criteria-group" style="border-color: #a78bfa;">
        <span class="criteria-group-label" style="color:#ffffff">기준점</span>
        <div class="criteria-input-wrapper">
           <label>Total</label>
           <input type="number" class="c-ref" placeholder="입력" value="${defaultRef}" step="0.1" style="width: 60px;">
        </div>
      </div>

      <div class="criteria-group">
        <span class="criteria-group-label" style="color:#a78bfa">상</span>
        <div class="criteria-input-wrapper">
           <label>최고</label>
           <input type="number" class="c-high-max" value="${hMax}" step="1">
        </div>
        <div class="criteria-input-wrapper">
           <label>최하</label>
           <input type="number" class="c-high-min" value="${hMin}" step="1">
        </div>
      </div>

      <div class="criteria-group">
        <span class="criteria-group-label" style="color:#646cff">중</span>
        <div class="criteria-input-wrapper">
           <label>최고</label>
           <input type="number" class="c-mid-max" value="${mMax}" step="1">
        </div>
        <div class="criteria-input-wrapper">
           <label>최하</label>
           <input type="number" class="c-mid-min" value="${mMin}" step="1">
        </div>
      </div>

      <div class="criteria-group">
        <span class="criteria-group-label" style="color:#2dd4bf">하</span>
        <div class="criteria-input-wrapper">
           <label>최고</label>
           <input type="number" class="c-low-max" value="${lMax}" step="1">
        </div>
        <div class="criteria-input-wrapper">
           <label>최하</label>
           <input type="number" class="c-low-min" value="${lMin}" step="1">
        </div>
      </div>
      <div class="criteria-input-wrapper" style="margin-left:auto">
         <label>간격</label>
         <input type="number" class="c-gap" value="${defaultGap}" step="0.1">
      </div>
    `;
        criteriaContainer.appendChild(row);

        const refInput = row.querySelector('.c-ref');
        const gapInput = row.querySelector('.c-gap');
        const hMaxInput = row.querySelector('.c-high-max');
        const hMinInput = row.querySelector('.c-high-min');
        const mMaxInput = row.querySelector('.c-mid-max');
        const mMinInput = row.querySelector('.c-mid-min');
        const lMaxInput = row.querySelector('.c-low-max');
        const lMinInput = row.querySelector('.c-low-min');

        // Master update function
        const applyCascadingConfig = (source) => {
            const gapVal = parseFloat(gapInput.value) || 0;
            let hMaxVal, hMinVal, mMaxVal, mMinVal;

            // 1. Determine High Values
            if (source === 'ref') {
                const val = parseFloat(refInput.value);
                if (!isNaN(val)) {
                    hMaxVal = Math.round(val);
                    hMinVal = Math.round(val);
                    hMaxInput.value = hMaxVal;
                    hMinInput.value = hMinVal;
                }
            } else if (source === 'mid') {
                // If source is Mid, we don't touch High, we just read Mid to update Low
            } else {
                // Gap or High input: Read High to drive Mid
                hMaxVal = parseFloat(hMaxInput.value);
                hMinVal = parseFloat(hMinInput.value);
            }

            // 2. Update Mid (if driven by High)
            if (source !== 'mid') {
                if (!isNaN(hMaxVal) && !isNaN(hMinVal)) {
                    // Mid Max = High Min - 1
                    const mMax = Math.round(hMinVal - 1);
                    // Mid Min = Mid Max - Gap
                    const mMin = Math.round(mMax - gapVal);
                    mMaxInput.value = mMax;
                    mMinInput.value = mMin;
                }
            }

            // 3. Update Low (driven by Mid)
            // Always read current Mid values (whether just updated or manually input)
            mMaxVal = parseFloat(mMaxInput.value);
            mMinVal = parseFloat(mMinInput.value);

            if (!isNaN(mMaxVal) && !isNaN(mMinVal)) {
                // Low Max = Mid Min - 1
                const lMax = Math.round(mMinVal - 1);
                // Low Min = Low Max - Gap
                const lMin = Math.round(lMax - gapVal);

                lMaxInput.value = lMax;
                lMinInput.value = lMin;
            }
        };

        refInput.addEventListener('input', () => applyCascadingConfig('ref'));
        hMaxInput.addEventListener('input', () => applyCascadingConfig('high'));
        hMinInput.addEventListener('input', () => applyCascadingConfig('high'));

        // Add listeners for Mid inputs to cascade to Low
        mMaxInput.addEventListener('input', () => applyCascadingConfig('mid'));
        mMinInput.addEventListener('input', () => applyCascadingConfig('mid'));

        gapInput.addEventListener('input', () => applyCascadingConfig('gap'));
    }
}

function generateGroupInputs() {
    const count = parseInt(document.getElementById('groupCount').value) || 0;

    if (count <= 0) {
        groupContainer.innerHTML = '<div class="empty-state">올바른 그룹 개수를 입력해주세요.</div>';
        return;
    }

    groupContainer.innerHTML = '';

    for (let i = 1; i <= count; i++) {
        const row = document.createElement('div');
        row.className = 'group-row';

        row.innerHTML = `
      <div class="group-label">그룹 ${i}</div>
      <div class="form-group">
        <label>상한선 (Max)</label>
        <input type="number" class="g-max" value="${100 - (i - 1) * 10}" step="0.1">
      </div>
      <div class="form-group">
        <label>하한선 (Min)</label>
        <input type="number" class="g-min" value="${90 - (i - 1) * 10}" step="0.1">
      </div>
    `;
        groupContainer.appendChild(row);
    }
}

function generateTable() {
    const entityCount = parseInt(document.getElementById('entityCount').value);
    // const rankScoreDiff = parseFloat(document.getElementById('rankScoreDiff').value); // Removed
    const judgeCount = parseInt(document.getElementById('judgeCount').value) || 1;
    const groupCount = parseInt(document.getElementById('groupCount').value);

    // Validation
    if (isNaN(entityCount) || isNaN(judgeCount)) {
        alert("설정값을 올바르게 입력해주세요.");
        return;
    }

    const groupRows = groupContainer.querySelectorAll('.group-row');
    if (groupRows.length === 0) {
        alert("먼저 '그룹 설정창 생성' 버튼을 눌러주세요.");
        return;
    }

    const groupSettings = Array.from(groupRows).map(row => ({
        max: parseFloat(row.querySelector('.g-max').value),
        min: parseFloat(row.querySelector('.g-min').value)
    }));

    const criteriaRows = criteriaContainer.querySelectorAll('.criteria-row');
    if (criteriaRows.length === 0) {
        alert("먼저 '기준 설정창 생성' 버튼을 눌러주세요.");
        return;
    }

    const criteriaSettings = Array.from(criteriaRows).map((row, i) => ({
        label: `기준 ${i + 1}`,
        high: { max: parseFloat(row.querySelector('.c-high-max').value), min: parseFloat(row.querySelector('.c-high-min').value) },
        mid: { max: parseFloat(row.querySelector('.c-mid-max').value), min: parseFloat(row.querySelector('.c-mid-min').value) },
        low: { max: parseFloat(row.querySelector('.c-low-max').value), min: parseFloat(row.querySelector('.c-low-min').value) },
        gap: parseFloat(row.querySelector('.c-gap').value)
    }));

    // Calculate Total Gap from all criteria
    const totalGap = criteriaSettings.reduce((sum, c) => sum + c.gap, 0);

    criteriaLabels = criteriaSettings.map(c => c.label);

    // 1. Generate Base Target Scores (Rank based)
    const baseEntities = [];

    // Exact Entity Distribution
    const baseCount = Math.floor(entityCount / groupSettings.length);
    let remainder = entityCount % groupSettings.length;

    let currentRank = 1;

    for (let g = 0; g < groupSettings.length; g++) {
        const group = groupSettings[g];

        let countInThisGroup = baseCount;
        if (remainder > 0) {
            countInThisGroup++;
            remainder--;
        }

        // Calculate step size
        // 1. Average Criteria Gap (User intention: Drop by about 1 grade-step per rank?)
        const avgGap = (criteriaSettings.length > 0) ? (totalGap / criteriaSettings.length) : 1.0;

        // 2. Density Step (Max sustainable step to fit everyone in range)
        const densityStep = (countInThisGroup > 1)
            ? (group.max - group.min) / (countInThisGroup - 1)
            : 0;

        // Use the smaller of the two to prevent clumping at the bottom,
        // while respecting the "Criteria Gap" as the standard step if space permits.
        const step = (densityStep > 0) ? Math.min(avgGap, densityStep) : 0;

        for (let k = 0; k < countInThisGroup; k++) {
            // Linear interpolation from Max
            let target = group.max - (step * k);

            // Add randomness (noise) relative to the step
            // e.g. +/- 30% of the step
            const noise = (step > 0) ? ((Math.random() * step * 0.6) - (step * 0.3)) : 0;
            target += noise;

            // Clamp
            if (target < group.min) target = group.min;
            if (target > group.max) target = group.max;

            // Round target to integer because criteria systems are integer-based
            const integerTarget = Math.round(target);
            baseEntities.push({
                rank: currentRank,
                targetScore: parseFloat(integerTarget.toFixed(1))
            });

            currentRank++;
        }
    }

    // 2. Generate Data for Each Judge
    globalJudgesData = [];

    for (let j = 0; j < judgeCount; j++) {
        const judgeEntities = baseEntities.map(base => {
            // Round target to integer because criteria are integers
            // This ensures Sum(Criteria) can generate the target
            const integerTarget = Math.round(base.targetScore);

            // Solve based on the requested target
            const scores = solveScoresInteger(integerTarget, criteriaSettings);

            // Calculate actual achievable total
            const total = scores.reduce((a, b) => a + b, 0);

            return {
                rank: base.rank,
                targetScore: total, // Actual achieved sum
                scores: scores,
                total: total
            };
        });
        globalJudgesData.push(judgeEntities);
    }

    // 3. Update UI
    // Populate Selector
    judgeSelector.innerHTML = '';
    for (let j = 0; j < judgeCount; j++) {
        const opt = document.createElement('option');
        opt.value = j;
        opt.innerText = `심사위원 ${j + 1}`;
        judgeSelector.appendChild(opt);
    }

    // Render Judge 1 (index 0)
    renderTable(globalJudgesData[0], criteriaSettings.length);
    document.getElementById('resultSection').classList.remove('hidden');
}

// Greedy Bucket Search + Strict Integer Solver
function solveScoresInteger(target, criteriaSettings) {
    const targetInt = Math.round(target);

    // 2. Global Bounds Check
    let absMax = 0;
    let absMin = 0;

    criteriaSettings.forEach(c => {
        absMax += Math.floor(c.high.max);
        absMin += Math.ceil(c.low.min);
    });

    let adjustedTarget = targetInt;
    if (adjustedTarget > absMax) adjustedTarget = absMax;
    if (adjustedTarget < absMin) adjustedTarget = absMin;

    // 3. Bucket Search
    let bucketState = criteriaSettings.map(() => 2); // 2=High, 1=Mid, 0=Low

    function getRange(state) {
        let min = 0, max = 0;
        state.forEach((s, idx) => {
            const c = criteriaSettings[idx];
            if (s === 2) { min += c.high.min; max += c.high.max; }
            else if (s === 1) { min += c.mid.min; max += c.mid.max; }
            else { min += c.low.min; max += c.low.max; }
        });
        return { min, max };
    }

    let range = getRange(bucketState);
    let iterations = 0;

    while ((adjustedTarget < range.min || adjustedTarget > range.max) && iterations < 1000) {
        iterations++;
        if (adjustedTarget < range.min) {
            const candidates = [];
            bucketState.forEach((s, i) => { if (s > 0) candidates.push(i); });
            if (candidates.length === 0) break;

            const maxLevel = Math.max(...bucketState);
            const bestCandidates = candidates.filter(i => bucketState[i] === maxLevel);
            const pick = bestCandidates[Math.floor(Math.random() * bestCandidates.length)];
            bucketState[pick]--;
        }
        else if (adjustedTarget > range.max) {
            const candidates = [];
            bucketState.forEach((s, i) => { if (s < 2) candidates.push(i); });
            if (candidates.length === 0) break;

            const minLevel = Math.min(...bucketState);
            const bestCandidates = candidates.filter(i => bucketState[i] === minLevel);
            const pick = bestCandidates[Math.floor(Math.random() * bestCandidates.length)];
            bucketState[pick]++;
        }
        range = getRange(bucketState);
    }

    // 4. Distribute
    const resultInts = [];
    const capacities = [];

    bucketState.forEach((s, idx) => {
        const c = criteriaSettings[idx];
        let p;
        if (s === 2) p = c.high;
        else if (s === 1) p = c.mid;
        else p = c.low;

        // Ensure we are selecting INTEGER values from the ranges
        const minVal = Math.ceil(p.min);
        const maxVal = Math.floor(p.max);

        resultInts.push(minVal);
        const cap = (maxVal > minVal) ? (maxVal - minVal) : 0;
        capacities.push({ idx, maxAdd: cap, currentAdd: 0 });
    });

    let currentSum = resultInts.reduce((a, b) => a + b, 0);
    let remainder = adjustedTarget - currentSum;

    let safety = 0;
    while (remainder > 0 && safety < 5000) {
        safety++;
        // Find candidates
        const candidates = capacities.filter(c => c.currentAdd < c.maxAdd);
        if (candidates.length === 0) break;

        // Balanced Distribution: Prioritize adding to the LOWEST scores first
        const currentScores = candidates.map(c => resultInts[c.idx]);
        const minScore = Math.min(...currentScores);

        // Only pick candidates that are currently at the minimum value
        const bestCandidates = candidates.filter(c => resultInts[c.idx] === minScore);
        const choice = bestCandidates[Math.floor(Math.random() * bestCandidates.length)];

        resultInts[choice.idx] += 1;
        choice.currentAdd += 1;
        remainder -= 1;
    }

    // 5. Post-Process: Smooth scores to enforce max difference <= 3 (where possible)
    const isValid = (val, c) => {
        // Check High
        if (val >= Math.ceil(c.high.min) && val <= Math.floor(c.high.max)) return true;
        // Check Mid
        if (val >= Math.ceil(c.mid.min) && val <= Math.floor(c.mid.max)) return true;
        // Check Low
        if (val >= Math.ceil(c.low.min) && val <= Math.floor(c.low.max)) return true;
        return false;
    };

    let pSafety = 0;
    // Monte Carlo Squashing: Try to reduce variance by swapping points from high scorers to low scorers
    while (pSafety < 5000) {
        pSafety++;

        // Find extreme candidates
        let maxIdx = -1, maxVal = -99999;
        let minIdx = -1, minVal = 99999;

        resultInts.forEach((v, i) => {
            if (v > maxVal) { maxVal = v; maxIdx = i; }
            if (v < minVal) { minVal = v; minIdx = i; }
        });

        if (maxVal - minVal <= 3) break; // Satisfied

        // Try to swap Max -> Min
        const nextMax = maxVal - 1;
        const nextMin = minVal + 1;

        if (isValid(nextMax, criteriaSettings[maxIdx]) && isValid(nextMin, criteriaSettings[minIdx])) {
            resultInts[maxIdx]--;
            resultInts[minIdx]++;
            continue; // Success, check again
        }

        // If strict Max->Min swap failed (due to range constraint), try semi-random swap
        // Pick a random High-ish and random Low-ish
        const highCandidates = [];
        const lowCandidates = [];
        const avg = adjustedTarget / criteriaSettings.length;

        resultInts.forEach((v, i) => {
            if (v > avg) highCandidates.push(i);
            if (v < avg) lowCandidates.push(i);
        });

        if (highCandidates.length === 0 || lowCandidates.length === 0) break;

        const i = highCandidates[Math.floor(Math.random() * highCandidates.length)];
        const j = lowCandidates[Math.floor(Math.random() * lowCandidates.length)];

        if (resultInts[i] - resultInts[j] > 3) {
            const nI = resultInts[i] - 1;
            const nJ = resultInts[j] + 1;
            if (isValid(nI, criteriaSettings[i]) && isValid(nJ, criteriaSettings[j])) {
                resultInts[i]--;
                resultInts[j]++;
            }
        }
    }

    return resultInts;
}

function renderTable(data, criteriaCount) {
    const table = document.getElementById('resultTable');
    const thead = table.querySelector('thead');
    const tbody = table.querySelector('tbody');

    let headerHTML = `<tr><th>등수</th><th>목표 합계</th>`;
    for (let i = 0; i < criteriaCount; i++) headerHTML += `<th>기준 ${i + 1}</th>`;
    headerHTML += `<th>검증(합계)</th></tr>`;
    thead.innerHTML = headerHTML;

    tbody.innerHTML = data.map(row => {
        // Highlight only if large mismatch (allow rounding diffs < 1.0)
        const diff = Math.abs(row.targetScore - row.total);
        const isMismatch = diff >= 1.0;
        const style = isMismatch ? 'color: #ff6b6b; font-weight:bold;' : '';

        let rowHTML = `<tr style="${style}">
          <td>${row.rank}</td>
          <td>${row.targetScore.toFixed(1)}</td>`;
        row.scores.forEach(s => rowHTML += `<td>${Number.isInteger(s) ? s : s.toFixed(1)}</td>`);
        rowHTML += `<td>${row.total.toFixed(1)}</td></tr>`;
        return rowHTML;
    }).join('');
}

function downloadExcel() {
    if (!globalJudgesData.length) return;

    const wb = utils.book_new();

    // Sheets for Judges
    globalJudgesData.forEach((judgeData, index) => {
        const wsData = [];
        const headers = ['등수', '목표점수', ...criteriaLabels, '합계', '오차검증'];
        wsData.push(headers);
        judgeData.forEach(row => {
            const diff = parseFloat((row.targetScore - row.total).toFixed(3));
            wsData.push([row.rank, row.targetScore, ...row.scores, row.total, diff === 0 ? 'OK' : diff]);
        });
        const ws = utils.aoa_to_sheet(wsData);
        utils.book_append_sheet(wb, ws, `심사위원 ${index + 1}`);
    });

    // Summary Sheet
    const summaryData = [];
    const judgeHeaders = globalJudgesData.map((_, i) => `심사위원 ${i + 1}`);
    const summaryHeader = ['등수', '이름(가제)', ...judgeHeaders, '총점 합계', '평균 점수'];
    summaryData.push(summaryHeader);

    const entityCount = globalJudgesData[0].length;
    for (let i = 0; i < entityCount; i++) {
        const row = [];
        const rank = globalJudgesData[0][i].rank;
        row.push(rank);
        row.push(`참가자 ${rank}`);

        let sumTotal = 0;
        for (let j = 0; j < globalJudgesData.length; j++) {
            const val = globalJudgesData[j][i].total;
            row.push(val);
            sumTotal += val;
        }

        const avg = parseFloat((sumTotal / globalJudgesData.length).toFixed(2));
        row.push(sumTotal.toFixed(2));
        row.push(avg);

        summaryData.push(row);
    }

    const wsSummary = utils.aoa_to_sheet(summaryData);
    utils.book_append_sheet(wb, wsSummary, '종합 결과');

    writeFile(wb, '심사결과_난수표.xlsx');
}
