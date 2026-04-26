document.addEventListener('DOMContentLoaded', () => {
    const uploadView = document.getElementById('upload-view');
    const appView = document.getElementById('app-view');
    const docxInput = document.getElementById('docx-input');
    const loadingState = document.getElementById('loading-state');
    const parseProgress = document.getElementById('parse-progress');
    const errorMsg = document.getElementById('upload-error');
    
    const questionGrid = document.getElementById('question-grid');
    const totalQ = document.getElementById('total-q');
    const listViewContent = document.getElementById('list-view-content');
    const questionViewContent = document.getElementById('question-view-content');
    const btnBackToList = document.getElementById('btn-back-to-list');
    const appSidebarTitle = document.getElementById('app-sidebar-title');
    
    const qTitle = document.getElementById('q-title');
    const qText = document.getElementById('q-text');
    const btnReveal = document.getElementById('btn-reveal');
    const answerContent = document.getElementById('answer-content');
    const qAnswerText = document.getElementById('q-answer-text');
    const qExplanationText = document.getElementById('q-explanation-text');
    const btnTranslate = document.getElementById('btn-translate');
    const translationLoading = document.getElementById('translation-loading');
    
    const btnNext = document.getElementById('btn-next');
    const btnPrev = document.getElementById('btn-prev');
    
    let parsedQuestions = [];
    let currentQuestionIndex = -1;

    // ----- DOCX Parsing using Mammoth.js -----
    docxInput.addEventListener('change', async (e) => {
        const file = e.target.files[0];
        if (!file) return;
        
        errorMsg.classList.add('hidden');
        
        // 拡張子を除外したファイル名をタイトルに設定
        const docTitle = file.name.replace(/\.[^/.]+$/, "");
        document.title = docTitle + ' - 試験問題学習アプリ';
        if (appSidebarTitle) {
            appSidebarTitle.textContent = docTitle;
        }

        loadingState.classList.remove('hidden');
        parseProgress.textContent = '読み込み中...';
        
        try {
            const arrayBuffer = await file.arrayBuffer();
            await processDOCX(arrayBuffer);
        } catch (error) {
            console.error(error);
            errorMsg.innerHTML = '解析に失敗しました:<br>' + error.message;
            errorMsg.classList.remove('hidden');
        } finally {
            loadingState.classList.add('hidden');
        }
    });

    async function processDOCX(arrayBuffer) {
        parseProgress.textContent = 'HTMLに変換中...';
        
        // Convert to HTML (mammoth natively extracts inline Base64 images)
        const result = await mammoth.convertToHtml({ arrayBuffer: arrayBuffer });
        const html = result.value;
        const messages = result.messages;
        if(messages && messages.length > 0) {
            console.warn("Mammoth messages:", messages);
        }
        
        parseProgress.textContent = '問題を構成中...';
        extractQuestions(html);
        
        if (parsedQuestions.length === 0) {
            throw new Error("問題の形式（'QUESTION NO: 1' 等）が見つかりませんでした。");
        }
        initAppUI();
    }

    function extractQuestions(htmlString) {
        parsedQuestions = [];
        const container = document.createElement('div');
        container.innerHTML = htmlString;
        
        // 1. Flatten the DOM to a sequential array of leaf-block elements
        // This strips away unpredictable wrappers like <table>, <ol>, <ul> 
        // that break our regex HTML string splitting.
        const flatElements = [];
        const blockTags = new Set(['P', 'LI', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'TD', 'DIV']);
        
        function traverseAndFlatten(node) {
            if (node.nodeType === 1) { // Element
                let hasBlockChild = false;
                for (let child of node.children) {
                    if (blockTags.has(child.tagName)) {
                        hasBlockChild = true;
                        break;
                    }
                }
                
                // If this is a block element and has no block children, it's a leaf block
                if (!hasBlockChild && blockTags.has(node.tagName)) {
                    flatElements.push(node);
                    return;
                }
                
                // Otherwise keep digging
                for (let child of node.childNodes) {
                    traverseAndFlatten(child);
                }
            } else if (node.nodeType === 3 && node.textContent.trim() !== '') {
                // If there's an orphan text node, treat it as a block
                if (node.parentNode === container || !blockTags.has(node.parentNode.tagName)) {
                    const p = document.createElement('p');
                    p.textContent = node.textContent;
                    flatElements.push(p);
                }
            }
        }
        
        traverseAndFlatten(container);
        
        let currentQuestion = null;
        let state = 'none'; // 'question', 'answer', 'explanation'
        
        const ansRegex = /\b(?:Answer|Correct\s+Answer)[:：]?\s*/i;
        const expRegex = /\b(?:Explanation|Explain|Reference)[:：]?\s*|={2,}/i;
        
        function appendToState(nodeOrHtml, targetState) {
            if (!currentQuestion) return;
            
            // Build the string to append based on whether it is an Element or a raw string from regex split
            let htmlToAppend = "";
            let elementHasImg = false;
            
            if (typeof nodeOrHtml === 'string') {
                if (nodeOrHtml.trim() === '') return;
                // Wrap in p tag if it's text slices
                htmlToAppend = `<div style="margin-bottom:0.5rem">${nodeOrHtml}</div>`;
            } else {
                if (nodeOrHtml.innerHTML.trim() === '' && !nodeOrHtml.querySelector('img')) return;
                elementHasImg = !!nodeOrHtml.querySelector('img');
                
                // Provide a safe wrapper
                const wrapper = document.createElement('div');
                wrapper.style.marginBottom = '0.5rem';
                wrapper.innerHTML = nodeOrHtml.innerHTML;
                htmlToAppend = wrapper.outerHTML;
            }
            
            if (targetState === 'question') currentQuestion.textHtml += htmlToAppend;
            else if (targetState === 'answer') currentQuestion.answerHtml += htmlToAppend;
            else if (targetState === 'explanation') currentQuestion.explanationHtml += htmlToAppend;
        }
        
        flatElements.forEach((el) => {
            const textContent = el.textContent.trim();
            const textUpper = textContent.toUpperCase();
            
            // Start of a new question
            const qMatch = textContent.match(/^QUESTION\s*(?:NO[:：]?)?\s*(\d+)/i);
            
            if (qMatch && !textUpper.includes('EXPLANATION')) {
                if (currentQuestion) parsedQuestions.push(currentQuestion);
                
                currentQuestion = {
                    id: qMatch[1],
                    textHtml: '',
                    answerHtml: '', 
                    explanationHtml: '',
                    plainExplanation: '', 
                    translatedExplanation: null
                };
                state = 'question';
                
                const cleanedHtml = el.innerHTML.replace(/^QUESTION\s*(?:NO[:：]?)?\s*\d+/i, '').trim();
                if (cleanedHtml || el.querySelector('img')) appendToState(cleanedHtml || el, 'question');
                return;
            }
            
            if (!currentQuestion) return; // Skip preamble
            
            const ansIdx = textContent.search(ansRegex);
            const expIdx = textContent.search(expRegex);
            
            const hasAns = ansIdx > -1;
            const hasExp = expIdx > -1;

            // Scenario 1: Both Answer and Explanation in same block
            if (hasAns && hasExp && ansIdx < expIdx) {
                // Because we flattened, el is just a leaf block (like <p>), so slicing innerHTML is totally safe!
                const qPart = el.innerHTML.slice(0, el.innerHTML.search(ansRegex)).trim();
                appendToState(qPart, state);
                
                state = 'explanation';
                
                const middleHtml = el.innerHTML.slice(el.innerHTML.search(ansRegex), el.innerHTML.search(expRegex));
                const ansPart = middleHtml.replace(ansRegex, '').trim();
                appendToState(ansPart, 'answer');
                
                const expPart = el.innerHTML.slice(el.innerHTML.search(expRegex)).replace(expRegex, '').trim();
                appendToState(expPart, 'explanation');
                currentQuestion.plainExplanation += textContent.slice(expIdx).replace(expRegex, '') + "\n";
                return;
            }
            
            // Scenario 2: Answer boundary
            if (hasAns && (!hasExp || expIdx < ansIdx)) {
                const prevPart = el.innerHTML.slice(0, el.innerHTML.search(ansRegex)).trim();
                appendToState(prevPart, state);
                
                state = 'answer';
                
                const ansPart = el.innerHTML.slice(el.innerHTML.search(ansRegex)).replace(ansRegex, '').trim();
                appendToState(ansPart, 'answer');
                return;
            }
            
            // Scenario 3: Explanation boundary
            if (hasExp) {
                const prevPart = el.innerHTML.slice(0, el.innerHTML.search(expRegex)).trim();
                appendToState(prevPart, state);
                
                state = 'explanation';
                
                const expPart = el.innerHTML.slice(el.innerHTML.search(expRegex)).replace(expRegex, '').trim();
                appendToState(expPart, 'explanation');
                currentQuestion.plainExplanation += textContent.slice(expIdx).replace(expRegex, '') + "\n";
                return;
            }
            
            // Typical line (No boundary)
            appendToState(el, state);
            if (state === 'explanation') {
                currentQuestion.plainExplanation += textContent + "\n";
            }
        });
        
        if (currentQuestion) parsedQuestions.push(currentQuestion);
        
        console.log(`Extracted ${parsedQuestions.length} exact questions.`);
    }

    // ----- UI Processing -----

    function initAppUI() {
        uploadView.classList.remove('active');
        setTimeout(() => uploadView.classList.add('hidden'), 300);
        appView.classList.remove('hidden');
        
        totalQ.textContent = parsedQuestions.length;
        
        questionGrid.innerHTML = '';
        parsedQuestions.forEach((q, idx) => {
            const btn = document.createElement('button');
            btn.className = 'q-btn';
            btn.textContent = q.id;
            btn.onclick = () => loadQuestion(idx);
            questionGrid.appendChild(btn);
        });
    }

    window.loadQuestion = function(index) {
        if (index < 0 || index >= parsedQuestions.length) return;
        currentQuestionIndex = index;
        const q = parsedQuestions[index];
        
        document.querySelectorAll('.q-btn').forEach((b, i) => {
            b.classList.toggle('active', i === index);
        });
        
        listViewContent.classList.remove('active');
        listViewContent.classList.add('hidden');
        btnBackToList.classList.remove('hidden');
        questionViewContent.classList.remove('hidden');
        questionViewContent.classList.add('active');
        
        qTitle.textContent = `Question ${q.id}`;
        
        // Insert parsed HTML directly (Mammoth safely escaped it and embedded Base64 images)
        qText.innerHTML = q.textHtml || '<p>問題文がありません</p>';
        
        btnReveal.classList.remove('hidden');
        answerContent.classList.add('hidden');
        qAnswerText.innerHTML = q.answerHtml || '<p>解答が見つかりませんでした</p>';
        qExplanationText.innerHTML = q.explanationHtml || '<p>解説がありません。</p>';
        
        btnTranslate.disabled = false;
        btnTranslate.textContent = '🇯🇵 日本語に翻訳';
        
        if (q.translatedExplanation) {
            qExplanationText.innerHTML = q.translatedExplanation;
            btnTranslate.classList.add('hidden');
        } else if (!q.plainExplanation || q.plainExplanation.trim() === '') {
            // Hide translate button if no text to translate
            btnTranslate.classList.add('hidden');
        } else {
            btnTranslate.classList.remove('hidden');
        }
        
        btnPrev.disabled = index === 0;
        btnNext.disabled = index === parsedQuestions.length - 1;
        
        setTimeout(() => {
            questionViewContent.style.opacity = 1;
        }, 50);
    }

    btnReveal.onclick = () => {
        btnReveal.classList.add('hidden');
        answerContent.classList.remove('hidden');
    };

    btnBackToList.onclick = () => {
        questionViewContent.classList.remove('active');
        questionViewContent.classList.add('hidden');
        btnBackToList.classList.add('hidden');
        listViewContent.classList.remove('hidden');
        listViewContent.classList.add('active');
        document.querySelectorAll('.q-btn').forEach(b => b.classList.remove('active'));
        currentQuestionIndex = -1;
    };

    btnPrev.onclick = () => loadQuestion(currentQuestionIndex - 1);
    btnNext.onclick = () => loadQuestion(currentQuestionIndex + 1);

    // ----- Translation (Google GTX API) -----

    btnTranslate.onclick = async () => {
        const q = parsedQuestions[currentQuestionIndex];
        if (!q.plainExplanation) return;
        
        btnTranslate.disabled = true;
        translationLoading.classList.remove('hidden');
        
        try {
            // Encode max 2000 chars per chunk ideally, but questions are usually short
            const textToTranslate = q.plainExplanation.slice(0, 3000); 
            const url = `https://translate.googleapis.com/translate_a/single?client=gtx&sl=en&tl=ja&dt=t&q=${encodeURIComponent(textToTranslate)}`;
            const res = await fetch(url);
            const data = await res.json();
            
            const translatedText = data[0].map(x => x[0]).join('');
            
            // To maintain images in explanation, we just append the translated text 
            // OR replace text nodes? Replacing text nodes is hard. 
            // Better: Prepend the translated text, keep original HTML below for images.
            const newHtml = `<div class="translated-text"><p>${translatedText.replace(/\n/g, '<br>')}</p></div><hr style="margin: 1rem 0; opacity: 0.2">` + q.explanationHtml;
            
            q.translatedExplanation = newHtml;
            qExplanationText.innerHTML = newHtml;
            btnTranslate.classList.add('hidden');
        } catch (e) {
            console.error('Translation error:', e);
            alert('翻訳に失敗しました。');
            btnTranslate.disabled = false;
        } finally {
            translationLoading.classList.add('hidden');
        }
    };
});
