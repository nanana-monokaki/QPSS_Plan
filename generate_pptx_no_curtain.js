const pptxgen = require("pptxgenjs");
const fs = require('fs');

async function createPresentation() {
    let pres = new pptxgen();
    pres.layout = 'LAYOUT_16x9'; // 10" x 5.625"
    pres.author = 'QPSS';
    pres.title = '聞き耳アワーシリーズ 企画書';

    // Theme Colors
    const COLOR_BG = "F7F7F7";       // Light Gray for background
    const COLOR_TEXT = "333333";     // Charcoal for text
    const COLOR_ACCENT = "C5283D";   // Surreal Red
    const COLOR_SUB_TEXT = "666666"; // Muted Gray
    const COLOR_CURTAIN = "1E2761";  // Deep Blue (Midnight Executive) for the curtain
    const COLOR_PANEL = "FFFFFF";

    // Master Slide Definition with Curtains
    pres.defineSlideMaster({
        title: 'MASTER_SLIDE',
        background: { path: "bg_texture.png" },
        objects: [
            // Slide Number
            { slideNumber: { x: 9.0, y: 5.35, w: 0.5, h: 0.2, fontSize: 10, color: COLOR_SUB_TEXT } }
        ]
    });

    // Helper functions
    const addSectionHeader = (slide, text) => {
        slide.addText(text, {
            x: 0.5, y: 0.2, w: 9.0, h: 0.8,
            fontSize: 32, bold: true, color: COLOR_CURTAIN,
            fontFace: "Yu Mincho",
            valign: "bottom"
        });
        slide.addShape(pres.shapes.LINE, {
            x: 0.5, y: 1.05, w: 2.0, h: 0,
            line: { color: COLOR_ACCENT, width: 3 }
        });
    };

    const addBodyText = (slide, textArr, x = 0.5, y = 1.6, w = 9.0, h = 3.0) => {
        slide.addText(textArr, {
            x: x, y: y, w: w, h: h,
            fontSize: 18, color: COLOR_TEXT,
            fontFace: "Yu Mincho",
            valign: "top",
            lineSpacing: 28,
            margin: 0
        });
    };

    // ==========================================
    // Slide 1: Title (P1)
    // ==========================================
    let slide1 = pres.addSlide({ masterName: "MASTER_SLIDE" });
    slide1.addText("直木賞作家・姫野カオルコプロデュース作品", {
        x: 0, y: 1.2, w: "100%", h: 0.6,
        fontSize: 16, color: COLOR_SUB_TEXT, align: "center", fontFace: "Yu Mincho"
    });
    slide1.addText("聞き耳アワーシリーズ", {
        x: 0, y: 1.8, w: "100%", h: 1.2,
        fontSize: 54, bold: true, color: COLOR_TEXT, align: "center", fontFace: "Yu Mincho",
        charSpacing: 4
    });
    // Add Placeholder Logo/Image for Title
    // Using an existing image if available, else a stylized shape
    slide1.addShape(pres.shapes.RECTANGLE, {
        x: 3.5, y: 3.2, w: 3.0, h: 1.5,
        fill: { color: COLOR_PANEL }, line: { color: COLOR_ACCENT, width: 2 },
        shadow: { type: "outer", color: "000000", blur: 5, offset: 3, angle: 45, opacity: 0.1 }
    });
    slide1.addText("Kikimimi Hour", {
        x: 3.5, y: 3.2, w: 3.0, h: 1.5,
        fontSize: 24, bold: true, color: COLOR_ACCENT, align: "center", fontFace: "Yu Mincho"
    });

    // ==========================================
    // Slide 2: 企画概要 (P2)
    // ==========================================
    let slide2 = pres.addSlide({ masterName: "MASTER_SLIDE" });
    addSectionHeader(slide2, "１．企画概要");

    slide2.addText("文学を「読む」から「劇場で聴く」へ。", {
        x: 0.5, y: 1.6, w: 9.0, h: 0.8,
        fontSize: 28, bold: true, color: COLOR_ACCENT, fontFace: "Yu Mincho"
    });

    slide2.addText("直木賞作家・姫野カオルコがプロデュース", {
        x: 0.5, y: 2.5, w: 9.0, h: 0.6,
        fontSize: 22, bold: true, color: COLOR_CURTAIN, fontFace: "Yu Mincho",
        fill: { color: "E2E8F0" }, align: "center"
    });

    addBodyText(slide2, [
        { text: "新しい文学体験型朗読劇シリーズです。", options: { breakLine: true } },
        { text: "選び抜かれた珠玉の物語を、劇場という空間で声・音・映像とともにQPSSが立体化。", options: { breakLine: true } },
        { text: "ライブ体験を起点とし、ポッドキャスト・音声配信へと展開、「耳から出会う文学」を日常へ届けます。" }
    ], 0.5, 3.4, 9.0, 1.5);

    // ==========================================
    // Slide 3: 企画意図 (P3)
    // ==========================================
    let slide3 = pres.addSlide({ masterName: "MASTER_SLIDE" });
    addSectionHeader(slide3, "２．企画意図");

    slide3.addText("「耳からはじまる文学体験」を劇場で創出させたい。", {
        x: 0.5, y: 1.5, w: 9.0, h: 0.5,
        fontSize: 20, bold: true, color: COLOR_TEXT, fontFace: "Yu Mincho"
    });

    // Left Panel: Reading Habits
    slide3.addText("活字離れの現状（1か月に本を1冊も読まない割合）", {
        x: 0.5, y: 2.1, w: 4.0, h: 0.4, fontSize: 14, bold: true, color: COLOR_TEXT, fontFace: "Yu Mincho", align: "center"
    });
    // Chart: 活字離れ (円グラフ)
    slide3.addChart(pres.charts.PIE, [{
        name: "読書状況",
        labels: ["読まない", "読む"],
        values: [62.6, 37.4]
    }], {
        x: 0.5, y: 2.5, w: 4.0, h: 2.3,
        showPercent: true,
        chartColors: [COLOR_ACCENT, "E2E8F0"],
        showLegend: true, legendPos: "b"
    });
    slide3.addText("出典：文化庁『国語に関する世論調査（令和5年度）』", {
        x: 0.5, y: 4.8, w: 4.0, h: 0.3, fontSize: 10, color: COLOR_SUB_TEXT, fontFace: "Yu Mincho", align: "center"
    });

    // Right Panel: Podcast Usage
    slide3.addText("国内ポッドキャスト月間利用率の推移", {
        x: 5.0, y: 2.1, w: 4.5, h: 0.4, fontSize: 14, bold: true, color: COLOR_TEXT, fontFace: "Yu Mincho", align: "center"
    });
    // Chart: Podcast (棒グラフ)
    slide3.addChart(pres.charts.BAR, [{
        name: "利用率",
        labels: ["2020年", "2021年", "2022年", "2023年", "2024年"],
        values: [14.2, 14.4, 15.7, 16.8, 17.2]
    }], {
        x: 5.0, y: 2.5, w: 4.5, h: 2.3, barDir: "col",
        chartColors: [COLOR_CURTAIN],
        showValue: true, dataLabelPosition: "outEnd",
        catGridLine: { style: "none" }, valGridLine: { color: "E2E8F0", size: 0.5 },
        showLegend: false
    });
    slide3.addText("出典：株式会社オトナル・株式会社朝日新聞社『ポッドキャスト国内利用実態調査』", {
        x: 5.0, y: 4.8, w: 4.5, h: 0.3, fontSize: 10, color: COLOR_SUB_TEXT, fontFace: "Yu Mincho", align: "center"
    });

    // Central Message
    slide3.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x: 2.0, y: 2.5, w: 6.0, h: 1.0,
        fill: { color: "FFFFFF", transparency: 10 },
        line: { color: COLOR_ACCENT, width: 2 },
        rectRadius: 0.1,
        shadow: { type: "outer", color: "000000", blur: 5, offset: 3, angle: 45, opacity: 0.15 }
    });
    slide3.addText("活字離れが進む一方で、「聴く時間はある」\nそんな現代人の生活リズムへ。", {
        x: 2.0, y: 2.5, w: 6.0, h: 1.0,
        fontSize: 18, bold: true, color: COLOR_TEXT, fontFace: "Yu Mincho", align: "center"
    });

    // ==========================================
    // Slide 4: 作品コンセプト (P4)
    // ==========================================
    let slide4 = pres.addSlide({ masterName: "MASTER_SLIDE" });
    addSectionHeader(slide4, "３．作品コンセプト");

    slide4.addText("少し怖くて、少し可笑しくて、なぜか心に残る物語たち。", {
        x: 0.5, y: 1.5, w: 9.0, h: 0.6,
        fontSize: 22, bold: true, color: COLOR_TEXT, fontFace: "Yu Mincho"
    });

    slide4.addShape(pres.shapes.RECTANGLE, {
        x: 0.5, y: 2.3, w: 4.0, h: 2.5, fill: { color: COLOR_PANEL }, line: { color: "E2E8F0", width: 1 }
    });
    slide4.addText([
        { text: "怪談、幻想譚、純文学、珠玉の短編", options: { bold: true, color: COLOR_ACCENT, breakLine: true } },
        { text: "——ジャンルを横断しながら、耳で聴くとより深く沁みる作品を姫野カオルコが選書。", options: { breakLine: true } },
        { text: " " },
        { text: "公演ごとに、違う世界へ迷い込むような感覚へいざなう連作型朗読劇シリーズ。", options: { breakLine: true } }
    ], { x: 0.7, y: 2.5, w: 3.6, h: 2.1, fontSize: 16, lineSpacing: 24, fontFace: "Yu Mincho" });

    slide4.addShape(pres.shapes.RECTANGLE, {
        x: 4.8, y: 2.3, w: 4.7, h: 2.5, fill: { color: COLOR_CURTAIN },
        shadow: { type: "outer", color: "000000", blur: 8, offset: 4, angle: 45, opacity: 0.2 }
    });
    slide4.addText([
        { text: "音楽・環境音による没入感ある演出の中、", options: { breakLine: true } },
        { text: "実力派キャストが朗読。", options: { breakLine: true, fontSize: 22, bold: true, color: COLOR_ACCENT } },
        { text: " " },
        { text: "ポッドキャスト等の「耳だけ」とは異なる、", options: { breakLine: true } },
        { text: "劇場空間ならではの【ライブ体験】", options: { breakLine: true, fontSize: 22, bold: true, color: "FFFFFF" } }
    ], { x: 5.0, y: 2.4, w: 4.3, h: 2.3, fontSize: 15, color: "E2E8F0", lineSpacing: 24, fontFace: "Yu Mincho" });

    // ==========================================
    // Slide 5: 姫野カオルコ (P5)
    // ==========================================
    let slide5 = pres.addSlide({ masterName: "MASTER_SLIDE" });
    addSectionHeader(slide5, "４．直木賞作家「姫野カオルコ」");

    slide5.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.6, w: 9.0, h: 3.4, fill: { color: COLOR_PANEL } });

    slide5.addText("キュレーターとして参加", {
        x: 0.5, y: 1.8, w: 4.0, h: 0.8,
        fontSize: 24, bold: true, color: COLOR_CURTAIN, fontFace: "Yu Mincho", align: "center",
        fill: { color: "E2E8F0" }
    });

    slide5.addText([
        { text: "鋭さとユーモア、そして人間の機微を描き続けてきた作家。", options: { breakLine: true, bold: true } },
        { text: " " },
        { text: "本シリーズでは以下を担い、キュレーターとして参加：", options: { breakLine: true } },
        { text: "・作品選定", options: { bullet: true, breakLine: true, indentLevel: 1 } },
        { text: "・世界観監修", options: { bullet: true, breakLine: true, indentLevel: 1 } },
        { text: "・文学的クオリティ統括", options: { bullet: true, breakLine: true, indentLevel: 1 } },
        { text: " " },
        { text: "単なる朗読劇ではなく、作家と一緒になって届ける文学レーベルとして展開します。", options: { breakLine: true, bold: true, color: COLOR_ACCENT } }
    ], { x: 4.8, y: 1.8, w: 4.5, h: 3.0, fontSize: 16, lineSpacing: 24, fontFace: "Yu Mincho" });

    // Placeholder for author photo
    slide5.addShape(pres.shapes.RECTANGLE, {
        x: 0.8, y: 2.8, w: 3.4, h: 2.0, fill: { color: "E2E8F0" }, line: { color: "CCCCCC", width: 1 }
    });
    slide5.addText("Author Image", {
        x: 0.8, y: 2.8, w: 3.4, h: 2.0, fontSize: 14, color: "999999", align: "center"
    });

    // ==========================================
    // Slide 6: ラジカル鈴木 (P6)
    // ==========================================
    let slide6 = pres.addSlide({ masterName: "MASTER_SLIDE" });
    addSectionHeader(slide6, "５．イラストレーター：ラジカル鈴木");

    slide6.addShape(pres.shapes.RECTANGLE, {
        x: 0.5, y: 1.6, w: 4.0, h: 3.4, fill: { color: COLOR_PANEL }, line: { color: COLOR_ACCENT, width: 2 }
    });
    slide6.addText([
        { text: "独特の線と色彩感覚で、", options: { breakLine: true } },
        { text: "どこか懐かしく、", options: { breakLine: true } },
        { text: "どこか不穏な幻想世界", options: { breakLine: true, bold: true, color: COLOR_ACCENT, fontSize: 22 } },
        { text: "を描くイラストレーター。", options: { breakLine: true } },
        { text: " " },
        { text: "ユーモアと毒気が共存する", options: { breakLine: true } },
        { text: "ビジュアルは、文学と深く響き合う。", options: { breakLine: true, bold: true } }
    ], { x: 0.8, y: 1.9, w: 3.4, h: 2.8, fontSize: 18, lineSpacing: 32, fontFace: "Yu Mincho" });

    // Placeholder for illustration
    slide6.addShape(pres.shapes.RECTANGLE, {
        x: 4.8, y: 1.6, w: 4.7, h: 3.4, fill: { color: "E2E8F0" }, line: { color: "CCCCCC", width: 1 }
    });
    slide6.addText("Illustration Image", {
        x: 4.8, y: 1.6, w: 4.7, h: 3.4, fontSize: 14, color: "999999", align: "center"
    });

    // ==========================================
    // Slide 7: QPSS (P7)
    // ==========================================
    let slide7 = pres.addSlide({ masterName: "MASTER_SLIDE" });
    addSectionHeader(slide7, "６．QPSSとは");

    slide7.addText("QUOBO PICTURES Screenwriters Studio（QPSS）", {
        x: 0.5, y: 1.5, w: 9.0, h: 0.5,
        fontSize: 18, bold: true, color: COLOR_TEXT, fontFace: "Yu Mincho"
    });

    slide7.addText("「物語の力で世界を動かす」知的生産工房。", {
        x: 0.5, y: 2.0, w: 9.0, h: 0.8,
        fontSize: 26, bold: true, color: COLOR_ACCENT, fontFace: "Yu Mincho", align: "center",
        fill: { color: "FFFFFF" },
        shadow: { type: "outer", color: "000000", blur: 4, offset: 2, angle: 45, opacity: 0.1 }
    });

    slide7.addText([
        { text: "映画・ドラマの企画開発から、ゲーム、企業キャラクター開発まで、構成力と創造性であらゆる物語を設計・構築するクリエイティブ集団です。", options: { breakLine: true } },
        { text: " " },
        { text: "本シリーズでは、脚本構成・舞台演出設計・音響設計を統括。", options: { breakLine: true, bold: true } },
        { text: "文学作品を“舞台言語”へ翻訳する役割を担います。", options: { breakLine: true, bold: true } },
        { text: " " },
        { text: "HP：http://www.quobo-pic.com", options: { breakLine: true, color: COLOR_CURTAIN, underline: true } }
    ], { x: 4.0, y: 3.0, w: 5.5, h: 2.0, fontSize: 16, lineSpacing: 24, fontFace: "Yu Mincho" });

    // Logo Image
    slide7.addImage({
        path: "mix.png",
        x: 0.5, y: 3.0, w: 3.0, h: 2.0,
        sizing: { type: 'contain' }
    });

    // ==========================================
    // Slide 8: メディア展開 & マスコット (P8)
    // ==========================================
    let slide8 = pres.addSlide({ masterName: "MASTER_SLIDE" });
    addSectionHeader(slide8, "７．メディア展開 ＆ キャラクター");

    slide8.addText("本公演は単発イベントではなく、継続的なシリーズIPとして展開。", {
        x: 0.5, y: 1.5, w: 9.0, h: 0.5,
        fontSize: 18, bold: true, color: COLOR_TEXT, fontFace: "Yu Mincho"
    });

    // Left Column: Media Expansion
    slide8.addShape(pres.shapes.RECTANGLE, {
        x: 0.5, y: 2.1, w: 4.5, h: 2.8, fill: { color: COLOR_PANEL }, line: { color: COLOR_CURTAIN, width: 2 }
    });
    slide8.addText("展開予定", { x: 0.5, y: 2.1, w: 4.5, h: 0.5, fontSize: 18, bold: true, color: "FFFFFF", fill: { color: COLOR_CURTAIN }, align: "center", fontFace: "Yu Mincho" });
    slide8.addText([
        { text: "劇場朗読公演", options: { bullet: true, breakLine: true } },
        { text: "ポッドキャスト配信 / 音声アーカイブ", options: { bullet: true, breakLine: true } },
        { text: "サブスク配信", options: { bullet: true, breakLine: true } },
        { text: "書籍コラボ", options: { bullet: true, breakLine: true } },
        { text: "グッズ展開 / オリジナルキャラクター", options: { bullet: true, breakLine: true, bold: true, color: COLOR_ACCENT } }
    ], { x: 0.7, y: 2.7, w: 4.2, h: 2.0, fontSize: 14, lineSpacing: 18, fontFace: "Yu Mincho" });
    slide8.addText("聴く文学ブランドとして成長させ、最終的に「映像化」を目指します。", {
        x: 0.5, y: 4.5, w: 4.5, h: 0.4, fontSize: 13, bold: true, align: "center", color: COLOR_TEXT, fontFace: "Yu Mincho"
    });

    // Right Column: Mascot
    slide8.addShape(pres.shapes.RECTANGLE, {
        x: 5.2, y: 2.1, w: 4.3, h: 2.8, fill: { color: "FFFFFF" }, line: { color: COLOR_ACCENT, width: 2 },
        shadow: { type: "outer", color: "000000", blur: 5, offset: 2, angle: 45, opacity: 0.1 }
    });
    slide8.addText("オリジナルマスコット", { x: 5.2, y: 2.1, w: 4.3, h: 0.5, fontSize: 18, bold: true, color: "FFFFFF", fill: { color: COLOR_ACCENT }, align: "center", fontFace: "Yu Mincho" });
    slide8.addText("耳を象ったキャラクター「キキミミズク（仮）」。\nグッズ化やキャラクタービジネスでも展開し、\n新規層の愛着を醸成します。", {
        x: 5.4, y: 2.7, w: 3.9, h: 1.0, fontSize: 13, lineSpacing: 22, fontFace: "Yu Mincho", align: "center"
    });
    slide8.addShape(pres.shapes.OVAL, {
        x: 6.35, y: 3.6, w: 2.0, h: 1.2, fill: { color: "E2E8F0" }
    });
    slide8.addText("Mascot Image", {
        x: 6.35, y: 3.6, w: 2.0, h: 1.2, fontSize: 12, color: "999999", align: "center"
    });


    // ==========================================
    // Slide 9: ターゲット (P9)
    // ==========================================
    let slide9 = pres.addSlide({ masterName: "MASTER_SLIDE" });
    addSectionHeader(slide9, "８．ターゲット");

    slide9.addText("活字から離れていた新規層にも多角的にアプローチします。", {
        x: 0.5, y: 1.5, w: 9.0, h: 0.5,
        fontSize: 18, bold: true, color: COLOR_TEXT, fontFace: "Yu Mincho"
    });

    // Target 1
    slide9.addShape(pres.shapes.RECTANGLE, {
        x: 0.5, y: 2.1, w: 4.3, h: 1.2, fill: { color: COLOR_PANEL }, line: { color: COLOR_CURTAIN, width: 1 },
        shadow: { type: "outer", color: "000000", blur: 4, offset: 2, angle: 45, opacity: 0.1 }
    });
    slide9.addText("文学・小説ファン", {
        x: 0.5, y: 2.1, w: 4.3, h: 0.4, fill: { color: COLOR_CURTAIN }, color: "FFFFFF", fontSize: 16, bold: true, align: "center", fontFace: "Yu Mincho"
    });
    slide9.addText("⇒ 「姫野カオルコ氏の選書・監修」でアピール", {
        x: 0.5, y: 2.6, w: 4.3, h: 0.6, fontSize: 16, bold: true, color: COLOR_ACCENT, align: "center", fontFace: "Yu Mincho"
    });

    // Target 2
    slide9.addShape(pres.shapes.RECTANGLE, {
        x: 5.2, y: 2.1, w: 4.3, h: 1.2, fill: { color: COLOR_PANEL }, line: { color: COLOR_CURTAIN, width: 1 },
        shadow: { type: "outer", color: "000000", blur: 4, offset: 2, angle: 45, opacity: 0.1 }
    });
    slide9.addText("若年層・カルチャー層", {
        x: 5.2, y: 2.1, w: 4.3, h: 0.4, fill: { color: COLOR_CURTAIN }, color: "FFFFFF", fontSize: 16, bold: true, align: "center", fontFace: "Yu Mincho"
    });
    slide9.addText("⇒ 「2.5次元キャストの起用」でアピール", {
        x: 5.2, y: 2.6, w: 4.3, h: 0.6, fontSize: 16, bold: true, color: COLOR_ACCENT, align: "center", fontFace: "Yu Mincho"
    });

    // Other targets
    slide9.addShape(pres.shapes.RECTANGLE, {
        x: 2.5, y: 3.5, w: 5.0, h: 1.5, fill: { color: "F8FAFC" }, line: { color: "E2E8F0", width: 1 }
    });
    slide9.addText([
        { text: "その他の周辺ターゲット", options: { bold: true, breakLine: true } },
        { text: "・ ポッドキャストリスナー", options: { bullet: true, breakLine: true } },
        { text: "・ ミステリー／怪談好き", options: { bullet: true, breakLine: true } },
        { text: "・ 舞台・朗読劇ファン", options: { bullet: true } }
    ], { x: 2.7, y: 3.6, w: 4.6, h: 1.3, fontSize: 16, lineSpacing: 22, fontFace: "Yu Mincho" });

    // ==========================================
    // Slide 10: 実施概要 & 公演予定 (P10 & P11)
    // ==========================================
    let slide10 = pres.addSlide({ masterName: "MASTER_SLIDE" });
    addSectionHeader(slide10, "９．実施概要 ＆ 公演予定");

    // Table for details
    let tableData = [
        [{ text: "上映期間", options: { fill: { color: COLOR_CURTAIN }, color: "FFFFFF", bold: true, fontFace: "Yu Mincho" } }, "３日間を予定（最終日はマチネのみ）"],
        [{ text: "上演時間", options: { fill: { color: COLOR_CURTAIN }, color: "FFFFFF", bold: true, fontFace: "Yu Mincho" } }, "約90分（トークショー込み）"],
        [{ text: "出演", options: { fill: { color: COLOR_CURTAIN }, color: "FFFFFF", bold: true, fontFace: "Yu Mincho" } }, "2〜3名（2.5次元俳優を想定）"],
        [{ text: "会場", options: { fill: { color: COLOR_CURTAIN }, color: "FFFFFF", bold: true, fontFace: "Yu Mincho" } }, "小規模劇場（100席想定）"],
        [{ text: "公演回数", options: { fill: { color: COLOR_CURTAIN }, color: "FFFFFF", bold: true, fontFace: "Yu Mincho" } }, "5回公演"],
        [{ text: "チケット代金", options: { fill: { color: COLOR_CURTAIN }, color: "FFFFFF", bold: true, fontFace: "Yu Mincho" } }, "5000円"],
        [{ text: "将来展望", options: { fill: { color: COLOR_CURTAIN }, color: "FFFFFF", bold: true, fontFace: "Yu Mincho" } }, "シリーズ化／定期開催（※詳細は調整可能）"]
    ];

    slide10.addTable(tableData, {
        x: 0.5, y: 1.5, w: 9.0, colW: [2.5, 6.5],
        border: { pt: 1, color: "E2E8F0" },
        fill: { color: COLOR_PANEL },
        fontSize: 14,
        fontFace: "Yu Mincho",
        valign: "middle"
    });

    // Upcoming schedule
    slide10.addShape(pres.shapes.RECTANGLE, {
        x: 0.5, y: 3.8, w: 9.0, h: 1.3, fill: { color: "F8FAFC" }, line: { color: COLOR_ACCENT, width: 2 }
    });
    slide10.addText("今後の公演予定", {
        x: 0.5, y: 3.8, w: 9.0, h: 0.4, fontSize: 16, bold: true, align: "center", color: COLOR_ACCENT, fontFace: "Yu Mincho"
    });
    slide10.addText([
        { text: "【第１弾】 「エンドレス・ラブ」（徳間書店）11月または12月想定", options: { breakLine: true, bold: true } },
        { text: "【第２弾】 「X博士」", options: { breakLine: true } },
        { text: "【第３弾】 「探偵物語」", options: {} }
    ], { x: 1.0, y: 4.2, w: 8.0, h: 0.8, fontSize: 14, lineSpacing: 18, fontFace: "Yu Mincho" });

    // Save
    await pres.writeFile({ fileName: "聞き耳アワーシリーズ企画書_改訂版_帯なし.pptx" });
    console.log("PPTX Generation Complete");
}

createPresentation().catch(err => {
    console.error(err);
});
