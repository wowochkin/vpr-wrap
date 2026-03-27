(function () {
  'use strict';

  const PLACEHOLDERS = [
    '__INSTRUCTION__',
    '__SENTENCES_BLOCK__',
    '__SOURCE_PARAGRAPH__',
    '__QUESTIONS_INTRO__',
    '__QUESTIONS_ITEMS__',
    '__HINT__',
  ];

  /** Встроенная копия шаблона — работает при открытии index.html через file:// */
  const DEFAULT_TEMPLATE = `<style>
    span.ans {  border-bottom:1px dotted #000088; cursor:pointer; }
    span.ans.marked { border-bottom:1px solid #880000; color:#880000; font-weight:bold; background-color:#ffff88; }
    </style>
    
    <div style="border:1px solid #000000;"><center><b>__INSTRUCTION__</b></center></div><br>
    
    <p class=MsoNormal>Прочитайте текст.</p><br>
    
     
    <div style="border:2px solid #996633; background:#fff9ed; padding:10px;margin:10px 0px;">
__SENTENCES_BLOCK__

__SOURCE_PARAGRAPH__
    </div><br>
    
    <p class="qq">__QUESTIONS_INTRO__</p>
__QUESTIONS_ITEMS__

    <p align="justify" style="font-size:14pt;text-indent:0px;margin:0px;padding:0px;color:#886600"><i>__HINT__
    </i></p>
`;

  function escapeHtml(s) {
    return String(s)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  const ASSIGN_EXT = /\.(txt|htm|html)$/i;

  function isAssignmentHtmlFile(name, rawHead) {
    const n = (name || '').toLowerCase();
    if (/\.(htm|html)$/i.test(n)) return true;
    const head = (rawHead || '').slice(0, 800);
    return /^\s*<\s*html[\s>]/i.test(head) || /^\s*<!DOCTYPE\s+html/i.test(head);
  }

  /**
   * Word в «Веб-странице» вставляет в разметку &lt;p&gt; реальные переводы строк между
   * переносами в исходном HTML; innerText целиком по секции даёт лишние \n внутри абзаца,
   * parseAssignmentText режет по строкам и buildSentencesBlock создаёт лишние &lt;p&gt;.
   */
  function plainTextOneParagraph(el) {
    let t = el.innerText != null ? el.innerText : el.textContent || '';
    t = String(t).replace(/\r\n/g, '\n').replace(/\r/g, '\n').replace(/\u00a0/g, ' ');
    return t.replace(/\s+/g, ' ').trim();
  }

  /** Экспорт Word «Веб-страница» → плоский текст для существующего парсера. */
  function stripHtmlToAssignmentPlain(html) {
    try {
      const doc = new DOMParser().parseFromString(html, 'text/html');
      const body = doc.body;
      if (!body) return html;
      body.querySelectorAll('script, style').forEach(function (el) {
        el.remove();
      });
      const wordRoot =
        body.querySelector('div.WordSection1') ||
        body.querySelector('div[class*="WordSection"]') ||
        null;
      const root = wordRoot || body;

      let text;
      if (wordRoot) {
        const lines = [];
        wordRoot.querySelectorAll('p').forEach(function (p) {
          const one = plainTextOneParagraph(p);
          if (one) lines.push(one);
        });
        text = lines.join('\n');
      } else {
        let t = root.innerText != null ? root.innerText : root.textContent || '';
        t = String(t).replace(/\r\n/g, '\n').replace(/\r/g, '\n').replace(/\u00a0/g, ' ');
        let prev;
        do {
          prev = t;
          t = t.replace(/([^\n])\n([^\n])/g, '$1 $2');
        } while (t !== prev);
        text = t
          .split('\n')
          .map(function (l) {
            return l.replace(/\s+$/, '');
          })
          .join('\n');
      }
      return text;
    } catch (e) {
      return html;
    }
  }

  function prepareAssignmentRaw(raw, fileName) {
    if (!isAssignmentHtmlFile(fileName, raw)) return raw;
    return stripHtmlToAssignmentPlain(raw);
  }

  function isVprOutputFragment(html) {
    const s = String(html).slice(0, 120000);
    return (
      /class\s*=\s*["']?ans["']?/i.test(s) &&
      /\bonclick\s*=\s*["']?\s*fp\s*\(\d+\)/i.test(s)
    );
  }

  /** Готовый фрагмент: отдельные p.qq с А./Б./В. — не пересобирать span (там может быть обрезанный текст). */
  function vprHasStructuredQuestions(doc) {
    const qq = doc.querySelectorAll('p.qq');
    if (qq.length < 4) return false;
    let ok = 0;
    for (let i = 1; i <= 3 && i < qq.length; i++) {
      const t = qq[i].textContent.replace(/\s+/g, ' ').trim();
      if (/^[АВБ]\.\s*/u.test(t)) ok++;
    }
    return ok >= 3;
  }

  const DEFAULT_HEADER_INSTRUCTION =
    'Прочитайте текст, ответьте на вопросы и выполните задания.';

  /** Только шапка и сущности — тело с span.ans не трогаем. */
  function polishVprFragmentHtml(html) {
    let s = String(html);
    s = s.replace(
      /<center>\s*<b>\s*<\/b>\s*<\/center>/gi,
      '<center><b>' + escapeHtmlApplyNbsp(DEFAULT_HEADER_INSTRUCTION) + '</b></center>'
    );
    s = s.replace(/&amp;nbsp;/gi, '&nbsp;');
    return applyNbsp(s.trim());
  }

  /** Разбор одной строки вида «…Выделите… А. … Б. … В. … Чтобы выбрать…» */
  function tryParseMergedQuestionsAndHint(text) {
    const t = String(text).replace(/\s+/g, ' ').trim();
    if (!t) return null;
    const hintIdx = t.search(/Чтобы выбрать нужное предложение/i);
    const main = hintIdx >= 0 ? t.slice(0, hintIdx).trim() : t;
    const hint = hintIdx >= 0 ? t.slice(hintIdx).trim() : '';
    const m = main.match(
      /^([\s\S]+?)\s+А\.\s+([\s\S]+?)\s+Б\.\s+([\s\S]+?)\s+В\.\s+([\s\S]+)$/u
    );
    if (!m) return null;
    return {
      intro: m[1].trim(),
      questions: ['А. ' + m[2].trim(), 'Б. ' + m[3].trim(), 'В. ' + m[4].trim()],
      hint: hint,
    };
  }

  /**
   * Уже собранный фрагмент ВПР → плоский текст для parseAssignmentText.
   * Чинит случай, когда вопросы попали в &lt;p align=justify&gt; и qq — только номер «71».
   */
  function reconstructPlainFromVprHtml(html) {
    if (!isVprOutputFragment(html)) return null;
    try {
      const doc = new DOMParser().parseFromString(html, 'text/html');
      let instruction = '';
      const cb = doc.querySelector('center b');
      if (cb) instruction = cb.textContent.replace(/\s+/g, ' ').trim();

      let box = null;
      doc.querySelectorAll('div').forEach(function (d) {
        const st = d.getAttribute('style') || '';
        if (/#996633|996633/.test(st)) box = d;
      });
      if (!box) return null;

      const bodyLines = [];
      let sourceLine = '';
      Array.prototype.forEach.call(box.children, function (p) {
        if (p.tagName !== 'P') return;
        const st = ((p.getAttribute('style') || '') + ' ' + (p.getAttribute('align') || '')).toLowerCase();
        if (st.indexOf('text-align:right') !== -1 || st.indexOf('text-align: right') !== -1) {
          sourceLine = p.textContent.replace(/\s+/g, ' ').trim();
          return;
        }
        const spans = p.querySelectorAll('span.ans');
        if (!spans.length) return;
        const parts = [];
        spans.forEach(function (sp) {
          const v = (sp.getAttribute('value') || '').trim();
          const txt = sp.textContent.replace(/\s+/g, ' ').trim();
          if (v) parts.push('(' + v + ') ' + txt);
        });
        if (parts.length) bodyLines.push(parts.join(' '));
      });

      if (!bodyLines.length) return null;

      if (!instruction) {
        instruction = DEFAULT_HEADER_INSTRUCTION;
      }

      const qqTexts = [];
      doc.querySelectorAll('p.qq').forEach(function (p) {
        const tx = p.textContent.replace(/\s+/g, ' ').trim();
        if (tx && !/^\d+$/.test(tx)) qqTexts.push(tx);
      });

      let questionsIntro = '';
      const questions = [];
      let hint = '';

      const hintP = doc.querySelector('p[align="justify"], p[align=justify]');
      const hintRaw = hintP ? hintP.textContent.replace(/\s+/g, ' ').trim() : '';
      const mergedFromHint = hintRaw ? tryParseMergedQuestionsAndHint(hintRaw) : null;

      if (mergedFromHint) {
        questionsIntro = mergedFromHint.intro;
        mergedFromHint.questions.forEach(function (q) {
          questions.push(q);
        });
        hint = mergedFromHint.hint;
      } else if (qqTexts.length) {
        questionsIntro = qqTexts[0];
        for (let qi = 1; qi < qqTexts.length; qi++) {
          questions.push(qqTexts[qi]);
        }
        hint = hintRaw;
      } else {
        hint = hintRaw;
        const again = hint ? tryParseMergedQuestionsAndHint(hint) : null;
        if (again) {
          questionsIntro = again.intro;
          again.questions.forEach(function (q) {
            questions.push(q);
          });
          hint = again.hint;
        }
      }

      const lines = [];
      lines.push(instruction);
      bodyLines.forEach(function (bl) {
        lines.push(bl);
      });
      lines.push(sourceLine);
      lines.push('');
      lines.push(questionsIntro);
      questions.forEach(function (q) {
        lines.push(q);
      });
      lines.push('');
      lines.push(hint);
      return lines.join('\n');
    } catch (e) {
      return null;
    }
  }

  function prepareInputPlain(raw, fileName) {
    if (isVprOutputFragment(raw)) {
      const rebuilt = reconstructPlainFromVprHtml(raw);
      if (rebuilt) return rebuilt;
    }
    return prepareAssignmentRaw(raw, fileName);
  }

  function blobMimeForOutputFilename(name) {
    const ext = (name.split('.').pop() || '').toLowerCase();
    if (ext === 'htm' || ext === 'html') return 'text/html;charset=utf-8';
    return 'text/plain;charset=utf-8';
  }

  /**
   * Все варианты неразрывного пробела в один символ U+00A0, чтобы правила с \\s срабатывали
   * и по «настоящим» пробелам, и по уже вставленным &nbsp;/&#160; в разметке.
   */
  function decodeNbspToUnicode(s) {
    return String(s)
      .replace(/&amp;#xA0;/gi, '\u00a0')
      .replace(/&amp;#160;/gi, '\u00a0')
      .replace(/&amp;nbsp;/gi, '\u00a0')
      .replace(/&#xA0;/gi, '\u00a0')
      .replace(/&#160;/g, '\u00a0')
      .replace(/&nbsp;/gi, '\u00a0')
      .replace(/\u202f|\u2007/g, '\u00a0');
  }

  /** Подстановка &nbsp; по образцу ВПР (всегда применяется к выходному HTML). */
  function applyNbsp(text) {
    let t = decodeNbspToUnicode(text);
    const pairs = [
      [/(\s|^)В\s/g, '$1В&nbsp;'],
      [/(\s|^)И\s/g, '$1И&nbsp;'],
      [/(\s|^)С\s/g, '$1С&nbsp;'],
      [/(\s|^)К\s/g, '$1К&nbsp;'],
      [/(\s|^)О\s/g, '$1О&nbsp;'],
      [/(\s|^)У\s/g, '$1У&nbsp;'],
      [/(\s|^)А\s/g, '$1А&nbsp;'],
      [/(\s|^)Но\s/g, '$1Но&nbsp;'],
      [/\sв\s/g, ' в&nbsp;'],
      [/\sво\s/g, ' во&nbsp;'],
      [/\sна\s/g, ' на&nbsp;'],
      [/\sпо\s/g, ' по&nbsp;'],
      [/\sиз\s/g, ' из&nbsp;'],
      [/\sза\s/g, ' за&nbsp;'],
      [/\sс\s/g, ' с&nbsp;'],
      [/\sи\s/g, ' и&nbsp;'],
      [/\sне\s/g, ' не&nbsp;'],
      [/\sот\s/g, ' от&nbsp;'],
      [/\sдо\s/g, ' до&nbsp;'],
      [/\sоб\s/g, ' об&nbsp;'],
      [/\sт\.\s*е\.\s/g, ' т.&nbsp;е.&nbsp;'],
    ];
    for (const [re, rep] of pairs) {
      t = t.replace(re, rep);
    }
    t = t.replace(/(\d)\s+(рубл|копеек|доллар|лет|года|году|века)/g, '$1&nbsp;$2');
    return t.replace(/\u00a0/g, '&nbsp;');
  }

  /** Сначала экранирование HTML, затем типографские &nbsp; (чтобы сущность &nbsp; не стала &amp;nbsp;). */
  function escapeHtmlApplyNbsp(s) {
    return applyNbsp(escapeHtml(String(s == null ? '' : s)));
  }

  function splitNumberedSentences(line) {
    const re = /\((\d+)\)\s*/g;
    const indices = [];
    let m;
    while ((m = re.exec(line)) !== null) {
      indices.push({ n: parseInt(m[1], 10), start: m.index, len: m[0].length });
    }
    const sentences = [];
    for (let j = 0; j < indices.length; j++) {
      const start = indices[j].start + indices[j].len;
      const end = j + 1 < indices.length ? indices[j + 1].start : line.length;
      const text = line.slice(start, end).trim();
      sentences.push({ n: indices[j].n, text });
    }
    return sentences;
  }

  /** Строка источника в скобках; не одна только нумерация «(17)» без текста после. */
  function isSourceLine(line) {
    const t = line.trim();
    if (!/^\([^)]+\)\s*$/.test(t)) return false;
    if (/^\(\d+\)\s*$/.test(t)) return false;
    return true;
  }

  function parseAssignmentText(raw) {
    const lines = raw.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
    const trimmedEnd = lines.map((l) => l.replace(/\s+$/, ''));
    while (trimmedEnd.length && trimmedEnd[trimmedEnd.length - 1] === '') trimmedEnd.pop();

    if (!trimmedEnd.length) throw new Error('Пустой файл');

    const instruction = trimmedEnd[0].trim();
    let i = 1;
    while (i < trimmedEnd.length && trimmedEnd[i].trim() === '') i++;

    const bodyLines = [];
    while (i < trimmedEnd.length) {
      const line = trimmedEnd[i];
      if (line.trim() === '') {
        i++;
        continue;
      }
      if (isSourceLine(line)) break;
      bodyLines.push(line.trim());
      i++;
    }

    const sourceLine = i < trimmedEnd.length ? trimmedEnd[i].trim() : '';
    if (sourceLine) i++;
    while (i < trimmedEnd.length && trimmedEnd[i].trim() === '') i++;
    while (i < trimmedEnd.length && /^\d+$/.test(trimmedEnd[i].trim())) i++;
    while (i < trimmedEnd.length && trimmedEnd[i].trim() === '') i++;

    let questionsIntro = '';
    if (i < trimmedEnd.length) {
      questionsIntro = trimmedEnd[i].trim();
      i++;
    }
    while (i < trimmedEnd.length && trimmedEnd[i].trim() === '') i++;

    let questions = [];
    while (i < trimmedEnd.length) {
      const line = trimmedEnd[i].trim();
      if (line === '') {
        i++;
        continue;
      }
      if (/^[А-ЯA-ZЁ]\.\s*/u.test(line)) {
        questions.push(line);
        i++;
      } else {
        break;
      }
    }

    const hintParts = [];
    while (i < trimmedEnd.length) {
      hintParts.push(trimmedEnd[i]);
      i++;
    }
    let hint = hintParts.join('\n').trim();

    if (questions.length === 0 && (questionsIntro || hint)) {
      const blob = [questionsIntro, hint].filter(Boolean).join('\n').replace(/\s+/g, ' ').trim();
      const spl = tryParseMergedQuestionsAndHint(blob);
      if (spl) {
        questionsIntro = spl.intro;
        questions = spl.questions;
        hint = spl.hint;
      }
    }

    if (
      instruction === 'Прочитайте текст.' &&
      bodyLines.length &&
      /^\(\d+\)/.test(bodyLines[0])
    ) {
      instruction = DEFAULT_HEADER_INSTRUCTION;
    }

    return {
      instruction,
      bodyLines,
      sourceLine,
      questionsIntro,
      questions,
      hint,
    };
  }

  function buildSentencesBlock(bodyLines) {
    const pStyle = 'text-indent:40px; line-height:1.5';
    const chunks = [];
    for (const line of bodyLines) {
      const sents = splitNumberedSentences(line);
      if (!sents.length) continue;
      const inner = sents
        .map(({ n, text }) => {
          const innerHtml = escapeHtmlApplyNbsp(text);
          return `(${n})<span class="ans" name="m0" value="${n}" onclick="fp(${n})">${innerHtml} </span>`;
        })
        .join(' ');
      chunks.push(`    <p class=MsoNormal style="${pStyle}">${inner}</p>`);
    }
    return chunks.join('\n');
  }

  function buildSourceParagraph(sourceLine) {
    if (!sourceLine) return '';
    return `    <p style="text-align: right;"><i>${escapeHtmlApplyNbsp(sourceLine)}</i></p>`;
  }

  function buildQuestionsItems(lines) {
    return lines
      .map((line) => {
        const m = line.match(/^([А-ЯA-ZЁ])\.(\s*)(.*)$/u);
        if (!m) {
          return `    <p class="qq">${escapeHtmlApplyNbsp(line)}</p>`;
        }
        const letter = m[1];
        const sp = m[2] || ' ';
        const rest = m[3];
        return `    <p class="qq"><b>${letter}.</b>${sp}${escapeHtmlApplyNbsp(rest)}</p>`;
      })
      .join('\n');
  }

  function fillTemplate(template, data) {
    let out = template;
    const map = {
      __INSTRUCTION__: escapeHtmlApplyNbsp(data.instruction),
      __SENTENCES_BLOCK__: data.sentencesBlock,
      __SOURCE_PARAGRAPH__: data.sourceParagraph,
      __QUESTIONS_INTRO__: escapeHtmlApplyNbsp(data.questionsIntro),
      __QUESTIONS_ITEMS__: data.questionsItems,
      __HINT__: escapeHtmlApplyNbsp(data.hint.replace(/\n+/g, ' ').trim()),
    };
    for (const [k, v] of Object.entries(map)) {
      if (!out.includes(k)) {
        /* допускаем опущенные необязательные блоки */
      }
      out = out.split(k).join(v);
    }
    return out;
  }

  /** class="qq" | class='qq' | class=qq (Word/KIM) */
  const RE_P_CLASS_QQ = '<p\\s+class\\s*=\\s*(?:"qq"|\'qq\'|qq)\\b[^>]*>';

  /**
   * Готовый «образец» с полным HTML (как Код_образец.txt) без __PLACEHOLDERS__ —
   * подменяем типовые блоки на плейсхолдеры.
   */
  function injectPlaceholdersIntoLegacyTemplate(tpl) {
    if (tpl.indexOf('__SENTENCES_BLOCK__') !== -1) return tpl;
    let t = tpl;
    t = t.replace(/<center>\s*<b>[\s\S]*?<\/b>\s*<\/center>/i, '<center><b>__INSTRUCTION__</b></center>');
    t = t.replace(
      /(<div[^>]*#996633[^>]*>)\s*([\s\S]*?)\s*(<\/div>\s*<br\s*\/?>)/i,
      function (_, open, _inner, close) {
        return open + '\n__SENTENCES_BLOCK__\n\n__SOURCE_PARAGRAPH__\n    ' + close;
      }
    );
    const reTwoQqThenHint = new RegExp(
      '(' +
        RE_P_CLASS_QQ +
        '[\\s\\S]*?<\\/p>)(\\s*' +
        RE_P_CLASS_QQ +
        '[\\s\\S]*?<\\/p>)+\\s*(?=<p[^>]*align\\s*=\\s*["\']justify["\'])',
      'i'
    );
    t = t.replace(
      reTwoQqThenHint,
      '<p class="qq">__QUESTIONS_INTRO__</p>\n__QUESTIONS_ITEMS__\n\n    '
    );
    t = t.replace(
      /<p\s+align\s*=\s*["']justify["'][^>]*>\s*<i>[\s\S]*?<\/i>\s*<\/p>/i,
      '<p align="justify" style="font-size:14pt;text-indent:0px;margin:0px;padding:0px;color:#886600"><i>__HINT__\n    </i></p>'
    );
    /* Задания 7.x и др.: после оранжевого блока идут таблицы/select без p align=justify — вырезаем хвост, ставим плейсхолдеры. */
    if (t.indexOf('__QUESTIONS_INTRO__') === -1 && t.indexOf('__SOURCE_PARAGRAPH__') !== -1) {
      t = t.replace(
        /(__SOURCE_PARAGRAPH__\s*<\/div>\s*<br\s*\/?>)([\s\S]*)$/i,
        function (_, head, rest) {
          if (rest.indexOf('__QUESTIONS_INTRO__') !== -1) {
            return head + rest;
          }
          return (
            head +
            '\n\n<p class="qq">__QUESTIONS_INTRO__</p>\n__QUESTIONS_ITEMS__\n\n    <p align="justify" style="font-size:14pt;text-indent:0px;margin:0px;padding:0px;color:#886600"><i>__HINT__\n    </i></p>\n'
          );
        }
      );
    }
    if (t.indexOf('__HINT__') === -1) {
      t =
        t.replace(/\s+$/, '') +
        '\n\n<p align="justify" style="font-size:14pt;text-indent:0px;margin:0px;padding:0px;color:#886600"><i>__HINT__\n    </i></p>\n';
    }
    return t;
  }

  function ensureTemplatePlaceholders(tpl) {
    if (PLACEHOLDERS.every(function (p) {
      return tpl.indexOf(p) !== -1;
    })) {
      return tpl;
    }
    return injectPlaceholdersIntoLegacyTemplate(tpl);
  }

  const DEFAULT_HINT_LLM =
    'Чтобы выбрать нужное предложение, щёлкните левой кнопкой мыши в любом его месте. После чего предложение выделится фоном. Чтобы отменить выбор, щёлкните повторно левой кнопкой мыши в любом месте этого предложения.';

  function buildOutputFromParsed(parsed, template, opts) {
    const sentencesBlock = buildSentencesBlock(parsed.bodyLines);
    const sourceParagraph = buildSourceParagraph(parsed.sourceLine);
    const questionsItems = buildQuestionsItems(parsed.questions);
    return fillTemplate(template, {
      instruction: parsed.instruction,
      sentencesBlock,
      sourceParagraph,
      questionsIntro: parsed.questionsIntro,
      questionsItems,
      hint: parsed.hint,
    });
  }

  function parseJsonFromLlmContent(content) {
    let s = String(content).trim();
    const fence = s.match(/```(?:json)?\s*([\s\S]*?)```/i);
    if (fence) s = fence[1].trim();
    return JSON.parse(s);
  }

  /** Нормализует ответ ИИ: в JSON приводим &nbsp;/U+00A0 к пробелу; после сборки HTML applyNbsp снова задаёт типографику. */
  function sanitizeLlmWhitespace(s) {
    if (s == null) return '';
    let t = String(s);
    t = t.replace(/&amp;nbsp;|&amp;#160;/gi, ' ');
    t = t.replace(/&nbsp;|&#160;|&#x0*a0;/gi, ' ');
    t = t.replace(/\u00a0/g, ' ');
    t = t.replace(/[ \t]{2,}/g, ' ');
    return t.trim();
  }

  function normalizeLlmParsed(obj) {
    let instruction = sanitizeLlmWhitespace(obj.instruction || '');
    if (!instruction) instruction = DEFAULT_HEADER_INSTRUCTION;
    let bodyLines = [];
    if (Array.isArray(obj.bodyLines)) {
      bodyLines = obj.bodyLines
        .map(function (x) {
          return sanitizeLlmWhitespace(x).replace(/\s+/g, ' ').trim();
        })
        .filter(Boolean);
    }
    if (!bodyLines.length && Array.isArray(obj.sentences)) {
      bodyLines = obj.sentences
        .map(function (x) {
          return sanitizeLlmWhitespace(x).replace(/\s+/g, ' ').trim();
        })
        .filter(Boolean);
    }
    const questions = Array.isArray(obj.questions)
      ? obj.questions
          .map(function (x) {
            return sanitizeLlmWhitespace(x);
          })
          .filter(Boolean)
      : [];
    let hint = sanitizeLlmWhitespace(obj.hint || '');
    if (!hint) hint = DEFAULT_HINT_LLM;
    return {
      instruction: instruction,
      bodyLines: bodyLines,
      sourceLine: sanitizeLlmWhitespace(obj.sourceLine || ''),
      questionsIntro: sanitizeLlmWhitespace(obj.questionsIntro || ''),
      questions: questions,
      hint: hint,
    };
  }

  /** У chat/completions DeepSeek допустимо только max_tokens ∈ [1, 8192]. */
  const DEEPSEEK_MAX_TOKENS = 8192;
  /** Верхняя граница длины user-текста (символы); дальше режем шаблон/задание поровну под бюджет. */
  const DEEPSEEK_USER_CONTENT_MAX_CHARS = 200000;

  /**
   * Читает SSE-поток chat/completions (stream: true).
   * @param {function(string): void} [onDelta] — накопленный полный текст ответа ассистента
   * @returns {{ fullText: string, usage: object|null }}
   */
  async function deepSeekReadStream(url, apiKey, bodyNoStream, onDelta) {
    const body = Object.assign({}, bodyNoStream, { stream: true });
    const res = await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: 'Bearer ' + apiKey,
      },
      body: JSON.stringify(body),
    });
    if (!res.ok) {
      const errText = await res.text();
      throw new Error('DeepSeek: ' + res.status + ' ' + errText.slice(0, 400));
    }
    if (!res.body || typeof res.body.getReader !== 'function') {
      throw new Error('DeepSeek: поток ответа недоступен (нет ReadableStream)');
    }
    const reader = res.body.getReader();
    const decoder = new TextDecoder();
    let lineBuf = '';
    let acc = '';
    let usage = null;
    for (;;) {
      const { done, value } = await reader.read();
      if (done) break;
      lineBuf += decoder.decode(value, { stream: true });
      let idx;
      while ((idx = lineBuf.indexOf('\n')) >= 0) {
        const line = lineBuf.slice(0, idx).replace(/\r$/, '').trim();
        lineBuf = lineBuf.slice(idx + 1);
        if (!line.startsWith('data:')) continue;
        const payload = line.slice(5).trim();
        if (payload === '[DONE]') continue;
        let j;
        try {
          j = JSON.parse(payload);
        } catch (e) {
          continue;
        }
        if (j.usage) usage = extractDeepSeekUsage(j);
        const ch = j.choices && j.choices[0];
        if (!ch) continue;
        const d = ch.delta;
        if (!d) continue;
        const part =
          d.content != null && d.content !== ''
            ? d.content
            : d.reasoning_content != null && d.reasoning_content !== ''
              ? d.reasoning_content
              : '';
        if (part) {
          acc += part;
          if (onDelta) onDelta(acc);
        }
      }
    }
    const tail = lineBuf.trim();
    if (tail.startsWith('data:')) {
      const payload = tail.slice(5).trim();
      if (payload !== '[DONE]') {
        try {
          const j = JSON.parse(payload);
          if (j.usage) usage = extractDeepSeekUsage(j);
          const d = j.choices && j.choices[0] && j.choices[0].delta;
          if (d && d.content) {
            acc += d.content;
            if (onDelta) onDelta(acc);
          }
        } catch (e) {}
      }
    }
    return { fullText: acc, usage: usage };
  }

  const DEEPSEEK_SYSTEM_PROMPT =
    'Ты извлекаешь структуру задания ВПР из сырого текста или HTML (Word/KIM, .txt, обрезанный фрагмент с span).\n' +
    'В запросе пользователя сначала даётся HTML-шаблон обвязки с плейсхолдерами (__INSTRUCTION__, __SENTENCES_BLOCK__, __SOURCE_PARAGRAPH__, __QUESTIONS_INTRO__, __QUESTIONS_ITEMS__, __HINT__) — это контекст целевой вёрстки; учитывай его при формулировках (например, стиль инструкции в шапке), но ответ всё равно только JSON, без HTML.\n' +
    'Верни один JSON-объект с ключами:\n' +
    '- instruction (string): фраза для шапки (если в источнике есть — возьми оттуда; иначе «Прочитайте текст, ответьте на вопросы и выполните задания.»).\n' +
    '- bodyLines (array of string): абзацы основного текста; в каждой строке сохрани нумерацию (1) (2) (3) … Полные предложения, без обрезки. Несколько номеров в одной строке — если так в источнике.\n' +
    '- sourceLine (string): строка вроде «(По …)» или пустая строка.\n' +
    '- questionsIntro (string): «Выделите в тексте предложения…» или аналог из источника.\n' +
    '- questions (array of string): три строки «А. …», «Б. …», «В. …» (кириллица А, Б, В).\n' +
    '- hint (string): подсказка про щелчок мыши; если в источнике нет — стандартная формулировка ВПР.\n' +
    'Не выдумывай факты; восстанавливай формулировки из источника. Игнорируй служебные номера в рамках (например «71»).\n' +
    'Число и текст вопросов в массиве questions должны соответствовать заданию (если не ровно три строки А/Б/В — отрази фактическую структуру задания, без «запасных» пунктов из шаблона).\n' +
    'Во всех строках JSON используй обычные пробелы; не вставляй HTML-сущности вроде &nbsp; и неразрывный пробел Unicode (U+00A0).';

  const DEEPSEEK_SYSTEM_PROMPT_FULL_HTML =
    'Ты готовишь итоговый HTML фрагмента учебного задания для веб-страницы.\n' +
    'В запросе пользователя два блока:\n' +
    '1) ОБРАЗЕЦ HTML — эталон вёрстки (таблицы, select, radio, div с рамкой #996633, p.qq, span.ans и onclick="fp(n)", классы, стили). Текст и число позиций в образце — пример; подставляй данные только из задания.\n' +
    '2) ЗАДАНИЕ — единственный источник истины по тексту и по количеству интерактивных элементов (строк таблицы, select, radio, нумерованных предложений).\n' +
    'Верни один готовый HTML (фрагмент или полный документ, как в образце). Требования:\n' +
    '- Повторяй вложенность тегов, классы, стили рамок, типы элементов (select, radio, span.ans) как в образце.\n' +
    '- КРИТИЧНО: все имена переменных и нумерация в разметке должны соответствовать ЗАДАНИЮ, а не копироваться слепо из образца. Атрибуты id, name, value, onclick="fp(n)", порядковые номера в таблице строк / вариантов вывода — вычисляй по фактическому содержанию задания (сколько пунктов, сколько вопросов, сколько предложений с номерами).\n' +
    '- Не допускай «лишних» полей из образца: если в задании меньше строк select/radio или меньше нумерованных предложений, чем в образце — убери лишнее; не оставляй id вроде a10 или name="a10", если в задании логика доходит только до a9 и т.п.\n' +
    '- Не допускай пропусков в последовательности: нумерация (n) в span.ans/fp(n), id a0,a1,… и связанные name должны идти непрерывно в рамках логики задания, без дыр и без «запасных» слотов с образца.\n' +
    '- Если в задании больше позиций, чем в образце — добавь строки/элементы по тому же шаблону разметки, с новыми последовательными id/name/value.\n' +
    '- Подставляй только факты и формулировки из задания; не выдумывай факты.\n' +
    '- СПИСКИ И ТАБЛИЦЫ ВАРИАНТОВ (понятия, ответы, строки «1)», «2)» в TABLE и т.п.): включай в HTML только то, что дословно или по нумерации есть в ЗАДАНИИ. Если в задании после вопроса нет перечня пунктов (нет таблицы/списка с вариантами) — не переноси таблицу из образца и не придумывай пункты «по смыслу»; оставь только текст из задания (ситуация + формулировка вопроса). Запрещено заполнять ячейки текстом из образца, если этих строк нет в задании.\n' +
    '- Для щелчка по предложению: span.ans, value и fp(n) должны совпадать с реальными номерами предложений в тексте задания, а не с числами из образца.\n' +
    '- Используй в разметке обычные HTML-сущности как в образце (&nbsp; и т.д.), не экранируй их как &amp;nbsp; внутри текста.\n' +
    '- Не добавляй markdown и не оборачивай ответ в ```; только HTML (при необходимости краткий комментарий <!-- -->).\n' +
    '- Игнорируй служебные номера задания в рамках вроде «72» в колонтитулах KIM.\n' +
    '- Перед ответом проверь: в HTML нет id/name/value/fp(n) из образца, которых нет в логике задания; нет пропусков в своей последовательности (например a0…a8 без a9, но с a10); нет таблицы/списка вариантов, которых не было в задании.';

  const DEEPSEEK_USER_REMINDER_VARS =
    '\n\n[ОБЯЗАТЕЛЬНО] Пересчитай все переменные только по этому заданию: id, name, value, onclick="fp(n)", число <tr>/<select>/<input type="radio">. Образец не задаёт «сколько полей» — только вид разметки. Удали лишние поля из образца; не копируй верхнюю границу индексов образца (пример: в образце был a10 — в твоём варианте последний индекс только если в задании есть десятая позиция). Нумерация подряд, без дыр.\n' +
    '[ОБЯЗАТЕЛЬНО] Таблица или список понятий/вариантов ответа: если в этом фрагменте задания их нет — не вставляй блок из образца и не выдумывай строки. Только текст, который реально есть в задании ниже.\n';

  /** Сборка user-сообщения: шаблон + задание, общий лимит ~maxPayload символов. */
  function buildDeepSeekUserContentFullHtml(rawText, fileName, templateExampleHtml, maxPayload) {
    const name = fileName || 'неизвестно';
    const headTpl =
      '--- ОБРАЗЕЦ HTML (структура и вёрстка; замени содержимое данными из задания) ---\n';
    const midSep =
      '\n\n--- ЗАДАНИЕ (исходный файл; источник истины по тексту и по количеству полей) ---\nИмя файла: ' +
      name +
      '\n\nПравило: id/name/value/fp(n), число строк таблиц и блоков — только из этого задания, подряд без пропусков; образец выше — лишь HTML-шаблон, не копируй из него лишние индексы (например a10), если в задании другая размерность.\n' +
      'Если в задании нет нумерованного перечня понятий/ответов после вопроса — в результате не должно быть такой таблицы/списка (ни из образца, ни вымышленной).\n\n---\n\n';
    const overhead =
      headTpl.length + midSep.length + DEEPSEEK_USER_REMINDER_VARS.length;
    const rawStr = String(rawText || '').replace(/\r\n/g, '\n').replace(/\r/g, '\n');
    let tpl = String(templateExampleHtml || '').replace(/\r\n/g, '\n').replace(/\r/g, '\n');
    const minRaw = 12000;
    const budget = maxPayload - overhead;
    let maxTplLen = budget - minRaw;
    if (maxTplLen < 3000) {
      maxTplLen = Math.max(0, budget - 4000);
    }
    if (tpl.length > maxTplLen) {
      tpl = tpl.slice(0, maxTplLen) + '\n\n[…образец обрезан по длине…]';
    }
    let maxRaw = budget - tpl.length;
    if (maxRaw < 1) maxRaw = 1;
    let slice = rawStr;
    if (slice.length > maxRaw) {
      slice = slice.slice(0, maxRaw) + '\n\n[…фрагмент задания обрезан для запроса…]';
    }
    return headTpl + tpl + midSep + DEEPSEEK_USER_REMINDER_VARS + slice;
  }

  function extractHtmlFromLlmContent(content) {
    let s = String(content || '').trim();
    const fence = s.match(/```(?:html|htm)?\s*([\s\S]*?)```/i);
    if (fence) s = fence[1].trim();
    return s.trim();
  }

  /** После модели: правим сущности, затем типографские &nbsp; (см. decodeNbsp внутри applyNbsp). */
  function polishLlmHtmlOutput(html) {
    let s = String(html || '').trim();
    s = s.replace(/&amp;nbsp;/gi, '&nbsp;');
    s = s.replace(/&amp;#160;/gi, '&#160;');
    s = s.replace(/&amp;#xA0;/gi, '&#xA0;');
    return applyNbsp(s);
  }

  /** Сборка user-сообщения: шаблон + задание, общий лимит ~maxPayload символов. */
  function buildDeepSeekUserContent(rawText, fileName, templateHtml, maxPayload) {
    const name = fileName || 'неизвестно';
    const headTpl =
      '--- ШАБЛОН HTML (плейсхолдеры подставляются на клиенте; верни только JSON полей) ---\n';
    const midSep = '\n\n--- ЗАДАНИЕ ---\nИмя файла: ' + name + '\n\n---\n\n';
    const overhead = headTpl.length + midSep.length;
    const rawStr = String(rawText || '').replace(/\r\n/g, '\n').replace(/\r/g, '\n');
    let tpl = String(templateHtml || '').replace(/\r\n/g, '\n').replace(/\r/g, '\n');
    const minRaw = 12000;
    const budget = maxPayload - overhead;
    let maxTplLen = budget - minRaw;
    if (maxTplLen < 3000) {
      maxTplLen = Math.max(0, budget - 4000);
    }
    if (tpl.length > maxTplLen) {
      tpl = tpl.slice(0, maxTplLen) + '\n\n[…шаблон обрезан по длине…]';
    }
    let maxRaw = budget - tpl.length;
    if (maxRaw < 1) maxRaw = 1;
    let slice = rawStr;
    if (slice.length > maxRaw) {
      slice = slice.slice(0, maxRaw) + '\n\n[…фрагмент задания обрезан для запроса…]';
    }
    return headTpl + tpl + midSep + slice;
  }

  async function deepSeekExtractParsed(rawText, fileName, apiKey, apiBase, templateHtml, onStreamDelta) {
    const base = String(apiBase || 'https://api.deepseek.com/v1').replace(/\/$/, '');
    const url = base + '/chat/completions';
    const userContent = buildDeepSeekUserContent(
      rawText,
      fileName,
      templateHtml,
      DEEPSEEK_USER_CONTENT_MAX_CHARS
    );

    const messages = [
      { role: 'system', content: DEEPSEEK_SYSTEM_PROMPT },
      { role: 'user', content: userContent },
    ];

    const bodyBase = {
      model: 'deepseek-chat',
      messages: messages,
      temperature: 0.2,
      max_tokens: DEEPSEEK_MAX_TOKENS,
      response_format: { type: 'json_object' },
    };

    if (onStreamDelta) {
      let streamBody = Object.assign({}, bodyBase);
      let streamOut;
      try {
        streamOut = await deepSeekReadStream(url, apiKey, streamBody, onStreamDelta);
      } catch (e1) {
        delete streamBody.response_format;
        streamOut = await deepSeekReadStream(url, apiKey, streamBody, onStreamDelta);
      }
      const msg = streamOut.fullText;
      if (!msg || !msg.trim()) throw new Error('DeepSeek: пустой ответ (стрим)');
      let obj;
      try {
        obj = parseJsonFromLlmContent(msg);
      } catch (e) {
        throw new Error('DeepSeek: не удалось разобрать JSON: ' + (e.message || e));
      }
      const parsed = normalizeLlmParsed(obj);
      if (!parsed.bodyLines.length) {
        throw new Error('DeepSeek: в ответе нет непустого массива bodyLines');
      }
      return { parsed: parsed, usage: streamOut.usage };
    }

    const body = Object.assign({}, bodyBase);
    let res = await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: 'Bearer ' + apiKey,
      },
      body: JSON.stringify(body),
    });

    let rawResp = await res.text();
    if (!res.ok && res.status === 400 && body.response_format) {
      delete body.response_format;
      res = await fetch(url, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          Authorization: 'Bearer ' + apiKey,
        },
        body: JSON.stringify(body),
      });
      rawResp = await res.text();
    }
    if (!res.ok) {
      throw new Error('DeepSeek: ' + res.status + ' ' + rawResp.slice(0, 400));
    }
    let data;
    try {
      data = JSON.parse(rawResp);
    } catch (e) {
      throw new Error('DeepSeek: ответ не JSON: ' + rawResp.slice(0, 200));
    }
    const msg =
      data.choices &&
      data.choices[0] &&
      data.choices[0].message &&
      data.choices[0].message.content;
    if (!msg) throw new Error('DeepSeek: пустой ответ choices');
    let obj;
    try {
      obj = parseJsonFromLlmContent(msg);
    } catch (e) {
      throw new Error('DeepSeek: не удалось разобрать JSON: ' + (e.message || e));
    }
    const parsed = normalizeLlmParsed(obj);
    if (!parsed.bodyLines.length) {
      throw new Error('DeepSeek: в ответе нет непустого массива bodyLines');
    }
    const usage = extractDeepSeekUsage(data);
    return { parsed: parsed, usage: usage };
  }

  async function deepSeekGenerateFullHtml(
    rawText,
    fileName,
    apiKey,
    apiBase,
    templateExampleHtml,
    onStreamDelta
  ) {
    const base = String(apiBase || 'https://api.deepseek.com/v1').replace(/\/$/, '');
    const url = base + '/chat/completions';
    const userContent = buildDeepSeekUserContentFullHtml(
      rawText,
      fileName,
      templateExampleHtml,
      DEEPSEEK_USER_CONTENT_MAX_CHARS
    );

    const bodyBase = {
      model: 'deepseek-chat',
      messages: [
        { role: 'system', content: DEEPSEEK_SYSTEM_PROMPT_FULL_HTML },
        { role: 'user', content: userContent },
      ],
      temperature: 0.2,
      max_tokens: DEEPSEEK_MAX_TOKENS,
    };

    if (onStreamDelta) {
      const streamOut = await deepSeekReadStream(url, apiKey, bodyBase, onStreamDelta);
      const msg = streamOut.fullText;
      if (!msg || !msg.trim()) throw new Error('DeepSeek: пустой ответ (стрим)');
      const html = polishLlmHtmlOutput(extractHtmlFromLlmContent(msg));
      if (html.length < 80) {
        throw new Error('DeepSeek: в ответе слишком мало HTML');
      }
      if (!/<[a-zA-Z!]/.test(html)) {
        throw new Error('DeepSeek: ответ не похож на HTML');
      }
      return { html: html, usage: streamOut.usage };
    }

    const res = await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: 'Bearer ' + apiKey,
      },
      body: JSON.stringify(bodyBase),
    });

    const rawResp = await res.text();
    if (!res.ok) {
      throw new Error('DeepSeek: ' + res.status + ' ' + rawResp.slice(0, 400));
    }
    let data;
    try {
      data = JSON.parse(rawResp);
    } catch (e) {
      throw new Error('DeepSeek: ответ не JSON: ' + rawResp.slice(0, 200));
    }
    const msg =
      data.choices &&
      data.choices[0] &&
      data.choices[0].message &&
      data.choices[0].message.content;
    if (!msg) throw new Error('DeepSeek: пустой ответ choices');
    const html = polishLlmHtmlOutput(extractHtmlFromLlmContent(msg));
    if (html.length < 80) {
      throw new Error('DeepSeek: в ответе слишком мало HTML');
    }
    if (!/<[a-zA-Z!]/.test(html)) {
      throw new Error('DeepSeek: ответ не похож на HTML');
    }
    const usage = extractDeepSeekUsage(data);
    return { html: html, usage: usage };
  }

  function extractDeepSeekUsage(data) {
    const u = data && data.usage;
    if (!u || typeof u !== 'object') return null;
    return {
      prompt_tokens: u.prompt_tokens,
      completion_tokens: u.completion_tokens,
      total_tokens: u.total_tokens,
      prompt_cache_hit_tokens: u.prompt_cache_hit_tokens,
      prompt_cache_miss_tokens: u.prompt_cache_miss_tokens,
    };
  }

  /** USD за 1M токенов (актуальные ориентиры DeepSeek, см. api-docs.deepseek.com). */
  const DEEPSEEK_USD_PER_1M_INPUT_CACHE_HIT = 0.028;
  const DEEPSEEK_USD_PER_1M_INPUT_CACHE_MISS = 0.28;
  const DEEPSEEK_USD_PER_1M_OUTPUT = 0.42;

  /**
   * Оценка стоимости запроса в USD.
   * Вход: cache_hit по дешёвому тарифу, cache_miss и остаток prompt (prompt − hit − miss) — по тарифу miss.
   * Выход: completion_tokens по тарифу output.
   */
  function estimateDeepSeekCostUsd(usage) {
    if (!usage || typeof usage !== 'object') return null;
    const hit = Number(usage.prompt_cache_hit_tokens) || 0;
    const miss = Number(usage.prompt_cache_miss_tokens) || 0;
    const prompt = Number(usage.prompt_tokens) || 0;
    const completion = Number(usage.completion_tokens) || 0;
    const restInput = Math.max(0, prompt - hit - miss);
    const inputUsd =
      (hit * DEEPSEEK_USD_PER_1M_INPUT_CACHE_HIT +
        (miss + restInput) * DEEPSEEK_USD_PER_1M_INPUT_CACHE_MISS) /
      1e6;
    const outputUsd = (completion * DEEPSEEK_USD_PER_1M_OUTPUT) / 1e6;
    const total = inputUsd + outputUsd;
    if (!isFinite(total) || (prompt === 0 && completion === 0)) {
      return null;
    }
    return total;
  }

  function formatDeepSeekCostUsd(usd) {
    if (usd == null || !isFinite(usd)) return '—';
    if (usd < 1e-4) return '~$' + usd.toFixed(6);
    if (usd < 0.01) return '~$' + usd.toFixed(5);
    return '~$' + usd.toFixed(4);
  }

  function formatUsageLine(u) {
    if (!u) return 'usage: (нет в ответе)';
    const h = u.prompt_cache_hit_tokens;
    const m = u.prompt_cache_miss_tokens;
    const p = u.prompt_tokens;
    const c = u.completion_tokens;
    const t = u.total_tokens;
    const cost = formatDeepSeekCostUsd(estimateDeepSeekCostUsd(u));
    return (
      'cache_hit ' +
      (h != null ? h : '—') +
      ' tok, cache_miss ' +
      (m != null ? m : '—') +
      ', prompt ' +
      (p != null ? p : '—') +
      ', completion ' +
      (c != null ? c : '—') +
      ', total ' +
      (t != null ? t : '—') +
      ', ≈ ' +
      cost
    );
  }

  async function processWithDeepSeek(
    rawText,
    templateWithPlaceholders,
    templateExampleUploaded,
    opts,
    sourceFileName,
    apiKey,
    apiBase,
    fullHtmlMode,
    onStreamDelta
  ) {
    if (fullHtmlMode) {
      const out = await deepSeekGenerateFullHtml(
        rawText,
        sourceFileName,
        apiKey,
        apiBase,
        templateExampleUploaded,
        onStreamDelta
      );
      return { html: out.html, usage: out.usage };
    }
    const out = await deepSeekExtractParsed(
      rawText,
      sourceFileName,
      apiKey,
      apiBase,
      templateWithPlaceholders,
      onStreamDelta
    );
    const html = buildOutputFromParsed(out.parsed, templateWithPlaceholders, opts);
    return { html: html, usage: out.usage };
  }

  function downloadBlob(blob, filename) {
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    a.click();
    URL.revokeObjectURL(a.href);
  }

  async function collectInputFiles(fileList) {
    const files = [];
    const arr = Array.from(fileList || []);
    for (const f of arr) {
      const name = f.name || '';
      if (name.toLowerCase().endsWith('.zip')) {
        if (typeof JSZip === 'undefined') throw new Error('JSZip не загружен');
        const z = await JSZip.loadAsync(f);
        const zipEntries = [];
        z.forEach((path, entry) => {
          if (entry.dir) return;
          if (path.includes('__MACOSX')) return;
          if (!ASSIGN_EXT.test(path)) return;
          zipEntries.push({ path, entry });
        });
        zipEntries.sort((a, b) => a.path.localeCompare(b.path, 'ru'));
        files.push(...zipEntries);
      } else if (ASSIGN_EXT.test(name)) {
        files.push({ path: name, file: f });
      }
    }
    return files;
  }

  async function readZipEntry(item) {
    return item.entry.async('string');
  }

  async function readFile(f) {
    return new Promise((resolve, reject) => {
      const r = new FileReader();
      r.onload = () => resolve(r.result);
      r.onerror = () => reject(r.error);
      r.readAsText(f, 'UTF-8');
    });
  }

  function bindDropzone(zone, input, opts) {
    const pick = opts.pick;
    const onReject = opts.onReject || function () {};

    function endDrag() {
      zone.classList.remove('dropzone--over');
    }

    zone.addEventListener('dragenter', (e) => {
      e.preventDefault();
      zone.classList.add('dropzone--over');
    });
    zone.addEventListener('dragleave', (e) => {
      e.preventDefault();
      const rel = e.relatedTarget;
      if (!rel || !zone.contains(rel)) endDrag();
    });
    zone.addEventListener('dragover', (e) => {
      e.preventDefault();
      e.dataTransfer.dropEffect = 'copy';
    });
    zone.addEventListener('drop', (e) => {
      e.preventDefault();
      endDrag();
      const files = Array.from(e.dataTransfer.files || []);
      const list = pick(files);
      if (!list.length) {
        onReject();
        return;
      }
      const dt = new DataTransfer();
      list.forEach(function (f) {
        dt.items.add(f);
      });
      input.files = dt.files;
      input.dispatchEvent(new Event('change', { bubbles: true }));
    });
  }

  function filesWordRu(n) {
    const m10 = n % 10;
    const m100 = n % 100;
    if (m100 >= 11 && m100 <= 14) return 'файлов';
    if (m10 === 1) return 'файл';
    if (m10 >= 2 && m10 <= 4) return 'файла';
    return 'файлов';
  }

  function formatTaskStatus(fileList) {
    if (!fileList || !fileList.length) return 'Не выбрано';
    const names = Array.from(fileList).map(function (f) {
      return f.name;
    });
    const n = names.length;
    if (n === 1) return '«' + names[0] + '»';
    if (n <= 3) return n + ' ' + filesWordRu(n) + ': ' + names.join(', ');
    return n + ' ' + filesWordRu(n) + ': ' + names.slice(0, 2).join(', ') + '…';
  }

  function init() {
    const elTemplate = document.getElementById('tplFile');
    const elTasks = document.getElementById('taskFiles');
    const elDropTpl = document.getElementById('dropzoneTpl');
    const elDropTasks = document.getElementById('dropzoneTasks');
    const elTplStatus = document.getElementById('tplStatus');
    const elTaskStatus = document.getElementById('taskStatus');
    const elDeepseekKey = document.getElementById('deepseekKey');
    const elDeepseekBase = document.getElementById('deepseekBase');
    const elRun = document.getElementById('btnRun');
    const elPreview = document.getElementById('preview');
    const elLog = document.getElementById('log');
    const elDownloadZip = document.getElementById('btnDownloadZip');
    const elDownloadOne = document.getElementById('btnDownloadOne');

    let deepSeekStreamEpoch = 0;
    let lastTemplate = '';
    let lastResults = [];
    let lastZipBlob = null;
    let lastSingleBlob = null;
    let lastSingleName = 'результат.txt';

    elDownloadZip.hidden = true;
    elDownloadOne.hidden = true;

    try {
      if (elDeepseekKey && localStorage.getItem('vpr-deepseek-key')) {
        elDeepseekKey.value = localStorage.getItem('vpr-deepseek-key');
      }
      if (elDeepseekBase && localStorage.getItem('vpr-deepseek-base')) {
        elDeepseekBase.value = localStorage.getItem('vpr-deepseek-base');
      } else if (elDeepseekBase && !elDeepseekBase.value) {
        elDeepseekBase.value = 'https://api.deepseek.com/v1';
      }
    } catch (e) {}

    lastTemplate = DEFAULT_TEMPLATE;
    elLog.textContent = '';
    fetch('templates/шаблон-vpr-по-умолчанию.txt')
      .then((r) => (r.ok ? r.text() : Promise.reject()))
      .then((t) => {
        lastTemplate = t;
        if (elTplStatus) elTplStatus.textContent = 'templates/шаблон-vpr-по-умолчанию.txt';
      })
      .catch(() => {});

    elTemplate.addEventListener('change', async () => {
      const f = elTemplate.files && elTemplate.files[0];
      if (!f) {
        if (elTplStatus) elTplStatus.textContent = 'Встроенный образец';
        return;
      }
      lastTemplate = await readFile(f);
      elLog.textContent = 'Шаблон: «' + f.name + '»';
      if (elTplStatus) elTplStatus.textContent = '«' + f.name + '»';
    });

    elTasks.addEventListener('change', () => {
      if (elTaskStatus) elTaskStatus.textContent = formatTaskStatus(elTasks.files);
    });

    if (elDropTpl && elTemplate) {
      bindDropzone(elDropTpl, elTemplate, {
        pick: function (files) {
          const txt = files.filter(function (f) {
            return (
              /\.(txt|htm|html)$/i.test(f.name) ||
              (f.type && (f.type.indexOf('text/') === 0 || f.type.indexOf('html') !== -1))
            );
          });
          return txt.length ? [txt[0]] : [];
        },
        onReject: function () {
          elLog.textContent = 'Нужен один файл .txt / .htm.';
        },
      });
    }

    if (elDropTasks && elTasks) {
      bindDropzone(elDropTasks, elTasks, {
        pick: function (files) {
          return files.filter(function (f) {
            return (
              ASSIGN_EXT.test(f.name) ||
              /\.zip$/i.test(f.name) ||
              (f.type &&
                (f.type.indexOf('zip') !== -1 ||
                  f.type.indexOf('text/') === 0 ||
                  f.type.indexOf('html') !== -1))
            );
          });
        },
        onReject: function () {
          elLog.textContent = 'Нужны .txt / .htm / .html или .zip.';
        },
      });
    }

    elRun.addEventListener('click', async () => {
      deepSeekStreamEpoch++;
      const streamEpochSnap = deepSeekStreamEpoch;
      elLog.textContent = '';
      elPreview.value = '';
      elDownloadZip.hidden = true;
      elDownloadOne.hidden = true;
      lastResults = [];
      lastZipBlob = null;
      lastSingleBlob = null;

      let template = lastTemplate;
      if (elTemplate.files && elTemplate.files[0]) {
        template = await readFile(elTemplate.files[0]);
      }
      if (!template || !template.trim()) {
        elLog.textContent = 'Нет образца HTML: загрузите .txt / .htm или откройте через сервер с templates/.';
        return;
      }

      const templateUploaded = template.trim();
      const opts = {};
      const useDeepSeekFullHtml = true;

      const templateWithPh = ensureTemplatePlaceholders(templateUploaded);
      const legacyAuto =
        templateWithPh !== templateUploaded &&
        templateUploaded.indexOf('__SENTENCES_BLOCK__') === -1;

      template = templateWithPh;
      const apiKey = elDeepseekKey ? elDeepseekKey.value.trim() : '';
      const apiBaseUrl =
        elDeepseekBase && elDeepseekBase.value.trim()
          ? elDeepseekBase.value.trim()
          : 'https://api.deepseek.com/v1';

      if (!apiKey) {
        elLog.textContent = 'Укажите API-ключ DeepSeek.';
        return;
      }
      try {
        localStorage.setItem('vpr-deepseek-key', apiKey);
        localStorage.setItem('vpr-deepseek-base', apiBaseUrl);
      } catch (e) {}

      const collected = await collectInputFiles(elTasks.files);
      if (!collected.length) {
        elLog.textContent = 'Выберите задания (.txt / .htm / .zip).';
        return;
      }

      const errors = [];
      const outputs = [];
      const prevBtnLabel = elRun.textContent;
      elRun.disabled = true;

      try {
        const n = collected.length;
        elLog.textContent =
          'DeepSeek · полный HTML · ' +
          n +
          ' файл(ов); первый — отдельный запрос' +
          (n > 1 ? ', остальные параллельно' : '') +
          '.';
        const results = [];
        const sum = { hit: 0, miss: 0, prompt: 0, completion: 0, total: 0, costUsd: 0 };

        function bumpUsage(u) {
          if (!u) return;
          sum.hit += Number(u.prompt_cache_hit_tokens) || 0;
          sum.miss += Number(u.prompt_cache_miss_tokens) || 0;
          sum.prompt += Number(u.prompt_tokens) || 0;
          sum.completion += Number(u.completion_tokens) || 0;
          sum.total += Number(u.total_tokens) || 0;
          const c = estimateDeepSeekCostUsd(u);
          if (c != null && isFinite(c)) sum.costUsd += c;
        }

        async function runDeepSeekItem(item, baseName, useStreamPreview) {
          let raw;
          try {
            if (item.file) raw = await readFile(item.file);
            else raw = await readZipEntry(item);
          } catch (e) {
            return { ok: false, name: baseName, err: 'не прочитан файл', usage: null };
          }
          try {
            const onStreamDelta =
              useStreamPreview && elPreview
                ? function (acc) {
                    if (streamEpochSnap !== deepSeekStreamEpoch) return;
                    const head = '«' + baseName + '» — стрим\n\n';
                    elPreview.value = head + acc;
                    elPreview.scrollTop = elPreview.scrollHeight;
                  }
                : null;
            const out = await processWithDeepSeek(
              raw,
              template,
              templateUploaded,
              opts,
              baseName,
              apiKey,
              apiBaseUrl,
              useDeepSeekFullHtml,
              onStreamDelta
            );
            return { ok: true, name: baseName, text: out.html, usage: out.usage };
          } catch (e) {
            return { ok: false, name: baseName, err: e.message || String(e), usage: null };
          }
        }

        function pushDsResult(r) {
          bumpUsage(r.usage);
          elLog.textContent += '\n  «' + r.name + '»: ' + formatUsageLine(r.usage);
          if (r.ok) results.push({ ok: true, name: r.name, text: r.text });
          else results.push({ ok: false, name: r.name, err: r.err });
        }

        const first = collected[0];
        const firstName = first.file ? first.file.name : first.path.split('/').pop();
        pushDsResult(await runDeepSeekItem(first, firstName, true));

        if (n > 1) {
          elLog.textContent += '\nОстальные файлы…';
          const batch = await Promise.all(
            collected.slice(1).map(function (item) {
              const baseName = item.file ? item.file.name : item.path.split('/').pop();
              return runDeepSeekItem(item, baseName, false);
            })
          );
          for (let bi = 0; bi < batch.length; bi++) {
            pushDsResult(batch[bi]);
          }
        }

        elLog.textContent +=
          '\nИтого: hit ' +
          sum.hit +
          ', miss ' +
          sum.miss +
          ', prompt ' +
          sum.prompt +
          ', completion ' +
          sum.completion +
          ', total ' +
          sum.total +
          ', ≈ ' +
          formatDeepSeekCostUsd(sum.costUsd > 0 ? sum.costUsd : null);

        for (let ri = 0; ri < results.length; ri++) {
          const r = results[ri];
          if (r.ok) outputs.push({ name: r.name, text: r.text });
          else errors.push(r.name + ': ' + r.err);
        }
      } finally {
        elRun.disabled = false;
        elRun.textContent = prevBtnLabel;
      }

      lastResults = outputs;

      if (outputs.length === 1) {
        lastSingleName = outputs[0].name;
        lastSingleBlob = new Blob([outputs[0].text], {
          type: blobMimeForOutputFilename(lastSingleName),
        });
        elPreview.value = outputs[0].text;
        elDownloadOne.hidden = false;
      } else if (outputs.length > 1) {
        elPreview.value = outputs.map((o) => '--- ' + o.name + ' ---\n').join('') + outputs[0].text.slice(0, 2000) + '\n…';
        if (typeof JSZip === 'undefined') {
          elLog.textContent +=
            (elLog.textContent ? '\n' : '') +
            'Несколько файлов: нужен JSZip для ZIP. Проверьте подключение скрипта с CDN.';
        } else {
          const zip = new JSZip();
          for (const o of outputs) zip.file(o.name, o.text);
          lastZipBlob = await zip.generateAsync({ type: 'blob' });
          elDownloadZip.hidden = false;
        }
      }

      if (errors.length) {
        elLog.textContent += (elLog.textContent ? '\n' : '') + errors.join('\n');
      }
      if (outputs.length) {
        let tail = 'Готово: ' + outputs.length + ' файл(ов).';
        if (legacyAuto) {
          tail = 'Образец без плейсхолдеров — разбор как полный HTML.\n' + tail;
        }
        elLog.textContent += (elLog.textContent ? '\n' : '') + tail;
      }
    });

    elDownloadOne.addEventListener('click', () => {
      if (lastSingleBlob) downloadBlob(lastSingleBlob, lastSingleName);
    });

    elDownloadZip.addEventListener('click', () => {
      if (lastZipBlob) downloadBlob(lastZipBlob, 'vpr-обвязка.zip');
    });
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
