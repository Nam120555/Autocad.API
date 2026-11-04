using System.Text;

namespace MyFirstProject.MoHinh
{
    /// <summary>
    /// Tạo file HTML tự chứa (self-contained) với Three.js để trực quan hóa mô hình BIM cốt thép 3D.
    /// Giao diện chuẩn BIM: 5 góc nhìn, sidebar thống kê & bảng khối lượng, bộ lọc đai/gối,
    /// thanh công cụ mặt cắt (clipping plane) tương tác, nhãn đá kê gối G1-G8.
    /// </summary>
    public static class HtmlViewerGenerator
    {
        public static string Generate(string jsonData, string projectName)
        {
            var sb = new StringBuilder();

            sb.AppendLine("<!DOCTYPE html>");
            sb.AppendLine("<html lang=\"vi\">");
            sb.AppendLine("<head>");
            sb.AppendLine("<meta charset=\"UTF-8\">");
            sb.AppendLine("<meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">");
            sb.AppendLine($"<title>{projectName} - Mô Hình BIM Cốt Thép 3D</title>");
            sb.AppendLine("<script src=\"https://cdnjs.cloudflare.com/ajax/libs/three.js/r128/three.min.js\"></script>");
            sb.AppendLine("<script src=\"https://cdn.jsdelivr.net/npm/three@0.128.0/examples/js/controls/OrbitControls.js\"></script>");
            sb.AppendLine(GetCssStyles());
            sb.AppendLine("</head>");
            sb.AppendLine("<body>");
            sb.AppendLine(GetHtmlLayout(projectName));
            sb.AppendLine("<script>");
            sb.AppendLine($"const REBAR_DATA = {jsonData};");
            sb.AppendLine(GetJavaScript());
            sb.AppendLine("</script>");
            sb.AppendLine("</body>");
            sb.AppendLine("</html>");

            return sb.ToString();
        }

        private static string GetCssStyles()
        {
            return @"
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

    * { margin: 0; padding: 0; box-sizing: border-box; }

    :root {
        --bg-main: #0c1017;
        --bg-sidebar: #131a26;
        --bg-card: rgba(255, 255, 255, 0.04);
        --bg-card-hover: rgba(59, 130, 246, 0.12);
        --border-color: rgba(255, 255, 255, 0.08);
        --text-primary: #f8fafc;
        --text-secondary: #94a3b8;
        --text-muted: #64748b;
        --accent-blue: #3b82f6;
        --accent-cyan: #06b6d4;
        --accent-emerald: #10b981;
    }

    body {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
        background: var(--bg-main);
        color: var(--text-primary);
        overflow: hidden;
        height: 100vh;
        width: 100vw;
        display: flex;
    }

    /* ===== SIDEBAR ===== */
    .sidebar {
        width: 320px;
        min-width: 320px;
        background: var(--bg-sidebar);
        border-right: 1px solid var(--border-color);
        display: flex;
        flex-direction: column;
        height: 100vh;
        z-index: 100;
        box-shadow: 4px 0 24px rgba(0, 0, 0, 0.4);
    }

    .sidebar-header {
        padding: 16px 18px 12px;
        border-bottom: 1px solid var(--border-color);
    }

    .project-title {
        font-size: 13px;
        font-weight: 700;
        color: #e2e8f0;
        line-height: 1.4;
        word-break: break-word;
    }

    /* Stats Cards */
    .stats-row {
        display: grid;
        grid-template-columns: 1fr 1fr;
        gap: 10px;
        padding: 12px 18px;
        border-bottom: 1px solid var(--border-color);
    }

    .stat-card {
        background: var(--bg-card);
        border: 1px solid var(--border-color);
        border-radius: 8px;
        padding: 10px 12px;
        text-align: left;
    }

    .stat-number {
        font-size: 24px;
        font-weight: 700;
        color: #ffffff;
        line-height: 1.1;
    }

    .stat-label {
        font-size: 11px;
        color: var(--text-secondary);
        margin-top: 4px;
    }

    /* Filters Row */
    .filters-row {
        display: flex;
        align-items: center;
        gap: 16px;
        padding: 10px 18px;
        border-bottom: 1px solid var(--border-color);
        font-size: 12px;
        color: #cbd5e1;
    }

    .filter-item {
        display: flex;
        align-items: center;
        gap: 6px;
        cursor: pointer;
        user-select: none;
    }

    .filter-item input[type=""checkbox""] {
        accent-color: var(--accent-blue);
        width: 15px;
        height: 15px;
        cursor: pointer;
    }

    /* Section Header */
    .section-header {
        display: flex;
        align-items: center;
        justify-content: space-between;
        padding: 10px 18px;
        border-bottom: 1px solid var(--border-color);
        background: rgba(0, 0, 0, 0.15);
    }

    .section-title {
        font-size: 12px;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        color: #94a3b8;
    }

    .btn-deselect {
        background: rgba(255, 255, 255, 0.08);
        border: 1px solid var(--border-color);
        color: #cbd5e1;
        padding: 3px 10px;
        border-radius: 4px;
        font-size: 11px;
        cursor: pointer;
        transition: all 0.15s;
    }

    .btn-deselect:hover {
        background: rgba(255, 255, 255, 0.16);
        color: #ffffff;
    }

    /* Table Container */
    .table-wrapper {
        flex: 1;
        overflow-y: auto;
        overflow-x: hidden;
    }

    .table-wrapper::-webkit-scrollbar {
        width: 5px;
    }
    .table-wrapper::-webkit-scrollbar-thumb {
        background: rgba(255, 255, 255, 0.15);
        border-radius: 4px;
    }

    .rebar-table {
        width: 100%;
        border-collapse: collapse;
        font-size: 11.5px;
    }

    .rebar-table th {
        position: sticky;
        top: 0;
        background: #101622;
        color: #94a3b8;
        font-weight: 600;
        text-align: right;
        padding: 8px 10px;
        border-bottom: 1px solid var(--border-color);
        z-index: 10;
    }

    .rebar-table th:first-child {
        text-align: left;
        padding-left: 14px;
    }

    .rebar-table td {
        padding: 6px 10px;
        text-align: right;
        color: #cbd5e1;
        border-bottom: 1px solid rgba(255, 255, 255, 0.04);
        cursor: pointer;
    }

    .rebar-table td:first-child {
        text-align: left;
        padding-left: 14px;
        font-weight: 600;
        color: #ffffff;
    }

    .rebar-table tr:hover td {
        background: var(--bg-card-hover);
        color: #ffffff;
    }

    .rebar-table tr.active td {
        background: rgba(59, 130, 246, 0.25);
        color: #ffffff;
        font-weight: 600;
    }

    .group-dot {
        display: inline-block;
        width: 8px;
        height: 8px;
        border-radius: 50%;
        margin-right: 6px;
        vertical-align: middle;
    }

    .table-footer {
        padding: 10px 18px;
        border-top: 1px solid var(--border-color);
        background: #0d121c;
        display: flex;
        justify-content: space-between;
        font-size: 12px;
        font-weight: 700;
        color: #ffffff;
    }

    .hint-bar {
        padding: 8px 18px;
        font-size: 10.5px;
        color: #64748b;
        background: #090c13;
        border-top: 1px solid rgba(255, 255, 255, 0.04);
    }

    /* ===== VIEWPORT CONTAINER ===== */
    .viewport-container {
        flex: 1;
        position: relative;
        height: 100vh;
        overflow: hidden;
    }

    #viewport {
        width: 100%;
        height: 100%;
    }

    /* Floating View Buttons (Top Left) */
    .floating-views {
        position: absolute;
        top: 16px;
        left: 20px;
        display: flex;
        gap: 6px;
        background: rgba(19, 26, 38, 0.85);
        backdrop-filter: blur(12px);
        padding: 4px;
        border-radius: 8px;
        border: 1px solid var(--border-color);
        z-index: 50;
        box-shadow: 0 4px 16px rgba(0, 0, 0, 0.3);
    }

    .view-btn {
        background: transparent;
        border: none;
        color: #94a3b8;
        padding: 6px 14px;
        border-radius: 6px;
        font-size: 12px;
        font-weight: 600;
        cursor: pointer;
        transition: all 0.15s;
    }

    .view-btn:hover {
        color: #ffffff;
        background: rgba(255, 255, 255, 0.08);
    }

    .view-btn.active {
        background: var(--accent-blue);
        color: #ffffff;
    }

    /* Floating Actions (Top Right) */
    .floating-actions {
        position: absolute;
        top: 16px;
        right: 20px;
        display: flex;
        gap: 8px;
        z-index: 50;
    }

    .action-btn {
        background: rgba(19, 26, 38, 0.85);
        backdrop-filter: blur(12px);
        border: 1px solid var(--border-color);
        color: #cbd5e1;
        padding: 7px 14px;
        border-radius: 6px;
        font-size: 12px;
        font-weight: 500;
        cursor: pointer;
        transition: all 0.15s;
        box-shadow: 0 4px 16px rgba(0, 0, 0, 0.3);
    }

    .action-btn:hover {
        background: rgba(255, 255, 255, 0.12);
        color: #ffffff;
    }

    /* Floating Clipping Bar (Bottom Left) */
    .floating-clip-bar {
        position: absolute;
        bottom: 20px;
        left: 20px;
        display: flex;
        align-items: center;
        gap: 12px;
        background: rgba(19, 26, 38, 0.85);
        backdrop-filter: blur(12px);
        padding: 8px 16px;
        border-radius: 8px;
        border: 1px solid var(--border-color);
        z-index: 50;
        box-shadow: 0 4px 16px rgba(0, 0, 0, 0.3);
        font-size: 12px;
        color: #cbd5e1;
    }

    .clip-select {
        background: #0d121c;
        border: 1px solid var(--border-color);
        color: #ffffff;
        padding: 4px 10px;
        border-radius: 4px;
        font-size: 12px;
        cursor: pointer;
        outline: none;
    }

    .clip-select option {
        background: #131a26;
        color: #ffffff;
    }

    .clip-checkbox-label {
        display: flex;
        align-items: center;
        gap: 6px;
        cursor: pointer;
        user-select: none;
    }

    .clip-checkbox-label input {
        accent-color: var(--accent-blue);
        cursor: pointer;
    }

    .clip-slider {
        width: 120px;
        accent-color: var(--accent-blue);
        cursor: pointer;
    }

    /* Pedestal Badges (G1 - G8) in 3D Space */
    .pedestal-badge {
        position: absolute;
        background: rgba(30, 41, 59, 0.9);
        border: 1px solid rgba(148, 163, 184, 0.35);
        color: #ffffff;
        padding: 4px 12px;
        border-radius: 6px;
        font-size: 12px;
        font-weight: 700;
        letter-spacing: 0.5px;
        pointer-events: none;
        z-index: 40;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.5);
        user-select: none;
        transform: translate(-50%, -50%);
        transition: opacity 0.2s;
    }

    /* Tooltip */
    .tooltip {
        position: fixed;
        background: rgba(15, 23, 42, 0.95);
        border: 1px solid var(--accent-blue);
        border-radius: 6px;
        padding: 10px 14px;
        font-size: 11.5px;
        color: #f8fafc;
        pointer-events: none;
        z-index: 1000;
        box-shadow: 0 8px 24px rgba(0, 0, 0, 0.6);
        display: none;
        min-width: 170px;
    }

    .tooltip-row {
        display: flex;
        justify-content: space-between;
        margin: 3px 0;
    }

    .tooltip-label { color: #94a3b8; }
    .tooltip-val { font-weight: 600; color: #ffffff; }

    /* Loading Overlay */
    .loading-overlay {
        position: fixed;
        inset: 0;
        background: #0c1017;
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        z-index: 9999;
        transition: opacity 0.4s ease;
    }
    .loading-overlay.hidden { opacity: 0; pointer-events: none; }

    .spinner {
        width: 44px;
        height: 44px;
        border: 3px solid rgba(255, 255, 255, 0.1);
        border-top-color: var(--accent-blue);
        border-radius: 50%;
        animation: spin 0.8s linear infinite;
    }
    @keyframes spin { to { transform: rotate(360deg); } }
</style>";
        }

        private static string GetHtmlLayout(string projectName)
        {
            return $@"
<!-- Loading -->
<div class=""loading-overlay"" id=""loadingOverlay"">
    <div class=""spinner""></div>
    <div style=""margin-top: 16px; font-size: 13px; color: #94a3b8;"">Đang tải mô hình BIM cốt thép...</div>
</div>

<!-- Tooltip -->
<div class=""tooltip"" id=""tooltip""></div>

<!-- SIDEBAR -->
<div class=""sidebar"">
    <div class=""sidebar-header"">
        <div class=""project-title"" id=""projectTitle"">{projectName}</div>
    </div>

    <!-- Stats -->
    <div class=""stats-row"">
        <div class=""stat-card"">
            <div class=""stat-number"" id=""statBars"">0</div>
            <div class=""stat-label"">Số thanh</div>
        </div>
        <div class=""stat-card"">
            <div class=""stat-number"" id=""statRatio"">0</div>
            <div class=""stat-label"">Hàm lượng thép (kg/m³)</div>
        </div>
    </div>

    <!-- Filters -->
    <div class=""filters-row"">
        <label class=""filter-item"">
            <input type=""checkbox"" id=""chkStirrup"" checked onchange=""toggleStirrups(this.checked)"">
            Đai
        </label>
        <label class=""filter-item"">
            <input type=""checkbox"" id=""chkPedestal"" checked onchange=""togglePedestals(this.checked)"">
            Đá kê / gối
        </label>
        <label class=""filter-item"">
            <input type=""checkbox"" id=""chkConcrete"" checked onchange=""toggleConcrete(this.checked)"">
            Bê tông
        </label>
    </div>

    <!-- Section Heading -->
    <div class=""section-header"">
        <span class=""section-title"" id=""memberSectionTitle"">MŨ</span>
        <button class=""btn-deselect"" onclick=""clearSelection()"">Bỏ chọn</button>
    </div>

    <!-- Table -->
    <div class=""table-wrapper"">
        <table class=""rebar-table"">
            <thead>
                <tr>
                    <th>Nhóm</th>
                    <th>Ø</th>
                    <th>L (mm)</th>
                    <th>KL (kg)</th>
                </tr>
            </thead>
            <tbody id=""tableBody""></tbody>
        </table>
    </div>

    <!-- Total Footer -->
    <div class=""table-footer"">
        <span>Tổng cộng:</span>
        <span id=""statTotalWeight"">0 kg</span>
    </div>

    <!-- Hint -->
    <div class=""hint-bar"">
        Chọn thanh trên mô hình để làm nổi
    </div>
</div>

<!-- VIEWPORT CONTAINER -->
<div class=""viewport-container"">
    <!-- Floating View Tabs -->
    <div class=""floating-views"">
        <button class=""view-btn active"" id=""btnView3d"" onclick=""switchView('3d')"">3D</button>
        <button class=""view-btn"" id=""btnViewFront"" onclick=""switchView('front')"">Mặt chính</button>
        <button class=""view-btn"" id=""btnViewSide"" onclick=""switchView('side')"">Mặt bên</button>
        <button class=""view-btn"" id=""btnViewPlan"" onclick=""switchView('plan')"">Mặt bằng</button>
        <button class=""view-btn"" id=""btnViewSection"" onclick=""switchView('section')"">Trục giao</button>
    </div>

    <!-- Floating Actions -->
    <div class=""floating-actions"">
        <button class=""action-btn"" onclick=""resetCamera()"">⟲ Đặt lại</button>
        <button class=""action-btn"" onclick=""exportScreenshot()"">📸 Lưu ảnh PNG</button>
    </div>

    <!-- 3D Canvas -->
    <div id=""viewport""></div>

    <!-- Floating Clipping Toolbar -->
    <div class=""floating-clip-bar"">
        <select class=""clip-select"" id=""clipSelect"" onchange=""onClipAxisChange(this.value)"">
            <option value=""none"">Không cắt</option>
            <option value=""x"">Mặt cắt dọc (X)</option>
            <option value=""y"">Mặt cắt đứng (Y)</option>
            <option value=""z"">Mặt cắt ngang (Z)</option>
        </select>
        <label class=""clip-checkbox-label"">
            <input type=""checkbox"" id=""clipInvert"" onchange=""onClipInvertChange(this.checked)"">
            Đảo phía
        </label>
        <input type=""range"" class=""clip-slider"" id=""clipSlider"" min=""0"" max=""100"" value=""50"" style=""display:none;"" oninput=""onClipSliderChange(this.value)"">
        <span id=""clipPosVal"" style=""display:none; font-size:11px; color:#94a3b8;"">50%</span>
    </div>

    <!-- Container for 3D Pedestal Badges -->
    <div id=""badgesContainer""></div>
</div>";
        }

        private static string GetJavaScript()
        {
            return @"
// ====================================================================
// THREE.JS REBAR BIM VIEWER — Professional Bridge & Civil Structure
// ====================================================================

let scene, camera, renderer, controls;
let rebarGroups = {};       // groupName -> THREE.Group
let concreteMesh = null;
let pedestalGroup = null;
let memberData = null;
let centerOffset = new THREE.Vector3();
let activeGroup = null;
let currentView = '3d';

// Clipping Planes
let clipPlane = null;
let clipAxis = 'none';
let clipInvert = false;
let clipPercent = 50;

// Badges coordinates in 3D
let pedestalBadges = [];

const raycaster = new THREE.Raycaster();
const mouse = new THREE.Vector2();

// ===== INIT =====
function init() {
    memberData = REBAR_DATA.members[0];

    // Populate Sidebar Info
    const fileName = REBAR_DATA.fileName ? REBAR_DATA.fileName.split(/[\\/]/).pop() : memberData.name;
    document.getElementById('projectTitle').textContent = `${fileName} - ${memberData.name}`;
    document.getElementById('memberSectionTitle').textContent = memberData.name || 'MŨ';

    document.getElementById('statBars').textContent = memberData.totalBars.toLocaleString('vi-VN');
    document.getElementById('statRatio').textContent = memberData.steelRatio.toLocaleString('vi-VN', { maximumFractionDigits: 0 });
    document.getElementById('statTotalWeight').textContent = memberData.totalWeight.toLocaleString('vi-VN', { maximumFractionDigits: 1 }) + ' kg';

    setupScene();
    buildConcrete();
    buildPedestals();
    buildRebars();
    buildTable();

    // Hide Loading
    setTimeout(() => {
        document.getElementById('loadingOverlay').classList.add('hidden');
    }, 400);

    // Events
    window.addEventListener('resize', onResize);
    document.getElementById('viewport').addEventListener('mousemove', onMouseMove);
    document.getElementById('viewport').addEventListener('click', onClick);

    animate();
}

// ===== SCENE SETUP =====
function setupScene() {
    const vp = document.getElementById('viewport');

    scene = new THREE.Scene();
    scene.background = new THREE.Color(0x0a0f18);

    camera = new THREE.PerspectiveCamera(40, vp.clientWidth / vp.clientHeight, 10, 200000);

    renderer = new THREE.WebGLRenderer({ antialias: true, alpha: true, preserveDrawingBuffer: true });
    renderer.setSize(vp.clientWidth, vp.clientHeight);
    renderer.setPixelRatio(Math.min(window.devicePixelRatio, 2));
    renderer.shadowMap.enabled = true;
    renderer.shadowMap.type = THREE.PCFSoftShadowMap;
    renderer.localClippingEnabled = true; // Enables section clipping planes!
    vp.appendChild(renderer.domElement);

    controls = new THREE.OrbitControls(camera, renderer.domElement);
    controls.enableDamping = true;
    controls.dampingFactor = 0.08;
    controls.rotateSpeed = 0.8;
    controls.zoomSpeed = 1.2;
    controls.panSpeed = 0.8;
    controls.maxDistance = 100000;

    // Lights
    const ambient = new THREE.AmbientLight(0xffffff, 0.7);
    scene.add(ambient);

    const dirLight1 = new THREE.DirectionalLight(0xffffff, 0.8);
    dirLight1.position.set(10000, 20000, 15000);
    dirLight1.castShadow = true;
    scene.add(dirLight1);

    const dirLight2 = new THREE.DirectionalLight(0x80bfff, 0.4);
    dirLight2.position.set(-10000, 5000, -10000);
    scene.add(dirLight2);

    const dirLight3 = new THREE.DirectionalLight(0xffffff, 0.3);
    dirLight3.position.set(0, -10000, 0);
    scene.add(dirLight3);

    // Center Offset: place center of member at (0, 0, 0)
    centerOffset.set(-memberData.width / 2, -memberData.height / 2, -memberData.depth / 2);

    // Initial camera position (3D view)
    resetCamera();
}

// ===== BUILD CONCRETE SHAPE (WITH XA MU CHAMFERS) =====
function buildConcrete() {
    const w = memberData.width;
    const h = memberData.height;
    const d = memberData.depth;

    concreteMesh = new THREE.Group();

    // Shape for Xa Mu (Chamfered cantilever ends)
    const shape = new THREE.Shape();
    const overhang = Math.min(w * 0.22, 2200); // Taper overhang length
    const taperH = h * 0.45; // Height at ends

    // Profile in X-Y plane:
    shape.moveTo(0, h);
    shape.lineTo(w, h);
    shape.lineTo(w, h - taperH);
    shape.lineTo(w - overhang, 0);
    shape.lineTo(overhang, 0);
    shape.lineTo(0, h - taperH);
    shape.closePath();

    const extrudeSettings = {
        steps: 1,
        depth: d,
        bevelEnabled: true,
        bevelThickness: 20,
        bevelSize: 20,
        bevelSegments: 2
    };

    const geom = new THREE.ExtrudeGeometry(shape, extrudeSettings);
    geom.computeVertexNormals();

    const mat = new THREE.MeshPhysicalMaterial({
        color: 0x93a3b8,
        transparent: true,
        opacity: 0.12,
        roughness: 0.7,
        metalness: 0.05,
        clearcoat: 0.1,
        side: THREE.DoubleSide,
        depthWrite: false
    });

    const mesh = new THREE.Mesh(geom, mat);
    concreteMesh.add(mesh);

    // Edges
    const edgesGeo = new THREE.EdgesGeometry(geom, 25);
    const edgesMat = new THREE.LineBasicMaterial({ color: 0x475569, opacity: 0.5, transparent: true });
    const edges = new THREE.LineSegments(edgesGeo, edgesMat);
    concreteMesh.add(edges);

    concreteMesh.position.copy(centerOffset);
    scene.add(concreteMesh);
}

// ===== BUILD PEDESTALS (ĐÁ KÊ GỐI G1 - G8) =====
function buildPedestals() {
    pedestalGroup = new THREE.Group();
    pedestalBadges = [];

    const w = memberData.width;
    const h = memberData.height;
    const d = memberData.depth;

    // 8 pedestals: 2 rows of 4
    const pedW = Math.min(700, w / 16);
    const pedH = 150; // 150mm height
    const pedD = Math.min(650, d / 4);

    const pedGeo = new THREE.BoxGeometry(pedW, pedH, pedD);
    const pedMat = new THREE.MeshPhysicalMaterial({
        color: 0x334155,
        roughness: 0.8,
        metalness: 0.1
    });

    const numCols = 4;
    const marginX = w * 0.12;
    const stepX = (w - 2 * marginX) / (numCols - 1);

    const zFront = d * 0.28;
    const zBack = d * 0.72;

    const badgesContainer = document.getElementById('badgesContainer');
    badgesContainer.innerHTML = '';

    let badgeIdx = 1;

    // Row 1 (Front: G1, G2, G3, G4)
    for (let c = 0; c < numCols; c++) {
        const x = marginX + c * stepX;
        const y = h + pedH / 2;
        const z = zFront;

        const mesh = new THREE.Mesh(pedGeo, pedMat);
        mesh.position.set(x, y, z);
        pedestalGroup.add(mesh);

        createPedestalBadge(`G${badgeIdx}`, x, y + pedH / 2 + 100, z);
        badgeIdx++;
    }

    // Row 2 (Back: G5, G6, G7, G8)
    for (let c = 0; c < numCols; c++) {
        const x = marginX + c * stepX;
        const y = h + pedH / 2;
        const z = zBack;

        const mesh = new THREE.Mesh(pedGeo, pedMat);
        mesh.position.set(x, y, z);
        pedestalGroup.add(mesh);

        createPedestalBadge(`G${badgeIdx}`, x, y + pedH / 2 + 100, z);
        badgeIdx++;
    }

    pedestalGroup.position.copy(centerOffset);
    scene.add(pedestalGroup);
}

function createPedestalBadge(text, worldX, worldY, worldZ) {
    const el = document.createElement('div');
    el.className = 'pedestal-badge';
    el.textContent = text;
    document.getElementById('badgesContainer').appendChild(el);

    pedestalBadges.push({
        element: el,
        pos: new THREE.Vector3(worldX, worldY, worldZ)
    });
}

function updatePedestalBadges() {
    if (!pedestalGroup || !pedestalGroup.visible) {
        pedestalBadges.forEach(b => b.element.style.display = 'none');
        return;
    }

    const vp = document.getElementById('viewport');
    const widthHalf = vp.clientWidth / 2;
    const heightHalf = vp.clientHeight / 2;

    pedestalBadges.forEach(b => {
        const p = b.pos.clone().add(centerOffset);
        p.project(camera);

        if (p.z > 1) { // Behind camera
            b.element.style.display = 'none';
            return;
        }

        b.element.style.display = 'block';
        const screenX = (p.x * widthHalf) + widthHalf;
        const screenY = -(p.y * heightHalf) + heightHalf;
        b.element.style.left = `${screenX}px`;
        b.element.style.top = `${screenY}px`;
    });
}

// ===== BUILD REBAR MESHES =====
function buildRebars() {
    rebarGroups = {};

    memberData.groups.forEach(group => {
        const groupObj = new THREE.Group();
        groupObj.name = group.name;

        const col = new THREE.Color(
            group.color[0] / 255,
            group.color[1] / 255,
            group.color[2] / 255
        );

        const mat = new THREE.MeshPhysicalMaterial({
            color: col,
            roughness: 0.4,
            metalness: 0.65,
            clearcoat: 0.35,
            clearcoatRoughness: 0.25
        });

        group.bars.forEach(bar => {
            if (bar.vertices.length < 2) return;

            const radius = Math.max(3, bar.diameter / 2);

            for (let i = 0; i < bar.vertices.length - 1; i++) {
                const p1 = new THREE.Vector3(bar.vertices[i][0], bar.vertices[i][1], bar.vertices[i][2]);
                const p2 = new THREE.Vector3(bar.vertices[i + 1][0], bar.vertices[i + 1][1], bar.vertices[i + 1][2]);
                const length = p1.distanceTo(p2);
                if (length < 1) continue;

                const segs = Math.max(6, Math.min(10, Math.round(radius)));
                const geo = new THREE.CylinderGeometry(radius, radius, length, segs);

                const mesh = new THREE.Mesh(geo, mat);
                mesh.castShadow = true;

                // Center point
                const mid = new THREE.Vector3().addVectors(p1, p2).multiplyScalar(0.5);
                mesh.position.copy(mid);

                // Align cylinder with segment vector
                const dir = new THREE.Vector3().subVectors(p2, p1).normalize();
                const quat = new THREE.Quaternion().setFromUnitVectors(new THREE.Vector3(0, 1, 0), dir);
                mesh.quaternion.copy(quat);

                mesh.userData = {
                    barId: bar.id,
                    groupName: group.name,
                    diameter: bar.diameter,
                    length: bar.length,
                    weight: bar.weight,
                    shapeType: bar.shapeType
                };

                groupObj.add(mesh);
            }

            // Elbow spherical joints
            if (bar.vertices.length > 2) {
                for (let i = 1; i < bar.vertices.length - 1; i++) {
                    const pt = bar.vertices[i];
                    const sphereGeo = new THREE.SphereGeometry(radius, 8, 8);
                    const sphere = new THREE.Mesh(sphereGeo, mat);
                    sphere.position.set(pt[0], pt[1], pt[2]);
                    groupObj.add(sphere);
                }
            }
        });

        groupObj.position.copy(centerOffset);
        scene.add(groupObj);
        rebarGroups[group.name] = groupObj;
    });
}

// ===== BUILD SIDEBAR TABLE =====
function buildTable() {
    const tbody = document.getElementById('tableBody');
    let html = '';

    memberData.groups.forEach(group => {
        const colStr = `rgb(${group.color[0]}, ${group.color[1]}, ${group.color[2]})`;
        const avgLen = group.bars.length > 0 ? group.bars[0].length : 0;

        html += `
            <tr id=""row-${group.name}"" onclick=""selectGroup('${group.name}')"">
                <td><span class=""group-dot"" style=""background: ${colStr};""></span>${group.name}</td>
                <td>${group.diameter}</td>
                <td>${avgLen.toLocaleString('vi-VN')}</td>
                <td>${group.totalWeight.toLocaleString('vi-VN', { minimumFractionDigits: 1, maximumFractionDigits: 1 })}</td>
            </tr>
        `;
    });

    tbody.innerHTML = html;
}

// ===== SELECT / HIGHLIGHT GROUP =====
function selectGroup(groupName) {
    if (activeGroup === groupName) {
        clearSelection();
        return;
    }

    activeGroup = groupName;

    // Update Table Active Row
    document.querySelectorAll('.rebar-table tr').forEach(r => r.classList.remove('active'));
    const targetRow = document.getElementById(`row-${groupName}`);
    if (targetRow) targetRow.classList.add('active');

    // Dim other groups, highlight active
    Object.entries(rebarGroups).forEach(([name, g]) => {
        const isMatch = (name === groupName);
        g.traverse(c => {
            if (c.material) {
                c.material.transparent = !isMatch;
                c.material.opacity = isMatch ? 1.0 : 0.12;
            }
        });
    });
}

function clearSelection() {
    activeGroup = null;
    document.querySelectorAll('.rebar-table tr').forEach(r => r.classList.remove('active'));

    Object.values(rebarGroups).forEach(g => {
        g.traverse(c => {
            if (c.material) {
                c.material.transparent = false;
                c.material.opacity = 1.0;
            }
        });
    });
}

// ===== FILTERS =====
function toggleStirrups(show) {
    Object.entries(rebarGroups).forEach(([name, g]) => {
        if (name.includes('Đai') || name === 'G3' || name.startsWith('ST')) {
            g.visible = show;
        }
    });
}

function togglePedestals(show) {
    if (pedestalGroup) pedestalGroup.visible = show;
    updatePedestalBadges();
}

function toggleConcrete(show) {
    if (concreteMesh) concreteMesh.visible = show;
}

// ===== VIEW SWITCHING =====
function switchView(view) {
    currentView = view;

    document.querySelectorAll('.view-btn').forEach(b => b.classList.remove('active'));

    const w = memberData.width;
    const h = memberData.height;
    const d = memberData.depth;
    const maxDim = Math.max(w, h, d);

    controls.target.set(0, 0, 0);

    switch(view) {
        case '3d':
            document.getElementById('btnView3d').classList.add('active');
            camera.position.set(w * 0.75, h * 1.5, d * 1.8);
            controls.enableRotate = true;
            break;
        case 'front':
            document.getElementById('btnViewFront').classList.add('active');
            camera.position.set(0, 0, maxDim * 1.6);
            controls.enableRotate = false;
            break;
        case 'side':
            document.getElementById('btnViewSide').classList.add('active');
            camera.position.set(maxDim * 1.6, 0, 0);
            controls.enableRotate = false;
            break;
        case 'plan':
            document.getElementById('btnViewPlan').classList.add('active');
            camera.position.set(0, maxDim * 1.6, 0.01);
            controls.enableRotate = false;
            break;
        case 'section':
            document.getElementById('btnViewSection').classList.add('active');
            camera.position.set(0, 0, maxDim * 1.5);
            controls.enableRotate = false;
            break;
    }

    camera.lookAt(0, 0, 0);
    controls.update();
}

function resetCamera() {
    switchView('3d');
}

// ===== CLIPPING PLANES (MẶT CẮT TÙY Ý) =====
function onClipAxisChange(axis) {
    clipAxis = axis;
    updateClippingPlane();
}

function onClipInvertChange(inverted) {
    clipInvert = inverted;
    updateClippingPlane();
}

function onClipSliderChange(percent) {
    clipPercent = parseFloat(percent);
    updateClippingPlane();
}

function updateClippingPlane() {
    const slider = document.getElementById('clipSlider');
    const posVal = document.getElementById('clipPosVal');

    if (clipAxis === 'none') {
        renderer.clippingPlanes = [];
        slider.style.display = 'none';
        posVal.style.display = 'none';
        return;
    }

    slider.style.display = 'inline-block';
    posVal.style.display = 'inline-block';
    posVal.textContent = `${clipPercent}%`;

    const w = memberData.width;
    const h = memberData.height;
    const d = memberData.depth;

    let normal = new THREE.Vector3();
    let maxDim = 1000;
    let minVal = 0;

    if (clipAxis === 'x') {
        normal.set(clipInvert ? 1 : -1, 0, 0);
        minVal = -w / 2;
        maxDim = w;
    } else if (clipAxis === 'y') {
        normal.set(0, clipInvert ? 1 : -1, 0);
        minVal = -h / 2;
        maxDim = h;
    } else if (clipAxis === 'z') {
        normal.set(0, 0, clipInvert ? 1 : -1);
        minVal = -d / 2;
        maxDim = d;
    }

    const currentPos = minVal + (clipPercent / 100) * maxDim;
    const constant = clipInvert ? -currentPos : currentPos;

    clipPlane = new THREE.Plane(normal, constant);
    renderer.clippingPlanes = [clipPlane];
}

// ===== INTERACTION & TOOLTIP =====
function onMouseMove(event) {
    const vp = document.getElementById('viewport');
    const rect = vp.getBoundingClientRect();
    mouse.x = ((event.clientX - rect.left) / rect.width) * 2 - 1;
    mouse.y = -((event.clientY - rect.top) / rect.height) * 2 + 1;

    raycaster.setFromCamera(mouse, camera);

    const meshes = [];
    Object.values(rebarGroups).forEach(g => {
        if (g.visible) {
            g.traverse(c => { if (c.isMesh) meshes.push(c); });
        }
    });

    const intersects = raycaster.intersectObjects(meshes);
    const tooltip = document.getElementById('tooltip');

    if (intersects.length > 0 && intersects[0].object.userData.barId) {
        const u = intersects[0].object.userData;
        tooltip.innerHTML = `
            <div class=""tooltip-row""><span class=""tooltip-label"">Mã thanh:</span><span class=""tooltip-val"">${u.barId}</span></div>
            <div class=""tooltip-row""><span class=""tooltip-label"">Nhóm:</span><span class=""tooltip-val"">${u.groupName}</span></div>
            <div class=""tooltip-row""><span class=""tooltip-label"">Đường kính:</span><span class=""tooltip-val"">Ø${u.diameter} mm</span></div>
            <div class=""tooltip-row""><span class=""tooltip-label"">Chiều dài:</span><span class=""tooltip-val"">${u.length.toLocaleString('vi-VN')} mm</span></div>
            <div class=""tooltip-row""><span class=""tooltip-label"">Khối lượng:</span><span class=""tooltip-val"">${u.weight} kg</span></div>
            <div class=""tooltip-row""><span class=""tooltip-label"">Dạng:</span><span class=""tooltip-val"">${u.shapeType}</span></div>
        `;
        tooltip.style.display = 'block';
        tooltip.style.left = `${event.clientX + 16}px`;
        tooltip.style.top = `${event.clientY - 10}px`;
    } else {
        tooltip.style.display = 'none';
    }
}

function onClick() {
    raycaster.setFromCamera(mouse, camera);
    const meshes = [];
    Object.values(rebarGroups).forEach(g => {
        if (g.visible) {
            g.traverse(c => { if (c.isMesh) meshes.push(c); });
        }
    });

    const intersects = raycaster.intersectObjects(meshes);
    if (intersects.length > 0 && intersects[0].object.userData.groupName) {
        selectGroup(intersects[0].object.userData.groupName);
    }
}

// ===== SCREENSHOT =====
function exportScreenshot() {
    renderer.render(scene, camera);
    const link = document.createElement('a');
    link.download = `${memberData.name || 'RebarBIM'}_3D.png`;
    link.href = renderer.domElement.toDataURL('image/png');
    link.click();
}

// ===== RESIZE =====
function onResize() {
    const vp = document.getElementById('viewport');
    camera.aspect = vp.clientWidth / vp.clientHeight;
    camera.updateProjectionMatrix();
    renderer.setSize(vp.clientWidth, vp.clientHeight);
    updatePedestalBadges();
}

// ===== ANIMATE LOOP =====
function animate() {
    requestAnimationFrame(animate);
    controls.update();
    updatePedestalBadges();
    renderer.render(scene, camera);
}

window.addEventListener('load', init);
";
        }
    }
}
