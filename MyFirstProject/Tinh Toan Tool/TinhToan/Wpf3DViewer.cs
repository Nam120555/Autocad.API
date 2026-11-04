using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Media3D;
using UserControl = System.Windows.Controls.UserControl;
using Point = System.Windows.Point;
using MouseEventArgs = System.Windows.Input.MouseEventArgs;
using Color = System.Windows.Media.Color;

namespace TinhToan
{
    public class Wpf3DViewer : UserControl
    {
        private Viewport3D _viewport;
        private PerspectiveCamera _camera;
        private ModelVisual3D _modelVisual;
        private Model3DGroup _mainGroup;

        // Orbit camera state
        private Point _lastMousePosition;
        private bool _isDragging = false;
        private double _cameraRadius = 150;
        private double _cameraTheta = Math.PI / 4;
        private double _cameraPhi = Math.PI / 4;

        public Wpf3DViewer()
        {
            this.Background = new SolidColorBrush(Color.FromRgb(25, 30, 40));
            _viewport = new Viewport3D();
            this.Content = _viewport;

            _camera = new PerspectiveCamera
            {
                FieldOfView = 60,
                UpDirection = new Vector3D(0, 0, 1)
            };
            _viewport.Camera = _camera;

            _mainGroup = new Model3DGroup();

            // Lighting
            _mainGroup.Children.Add(new AmbientLight(Color.FromRgb(120, 120, 120)));
            _mainGroup.Children.Add(new DirectionalLight(Color.FromRgb(180, 180, 180), new Vector3D(-1, -1, -1)));
            _mainGroup.Children.Add(new DirectionalLight(Color.FromRgb(80, 80, 90), new Vector3D(1, 1, -0.2)));

            _modelVisual = new ModelVisual3D { Content = _mainGroup };
            _viewport.Children.Add(_modelVisual);

            UpdateCameraPosition();

            this.MouseLeftButtonDown += OnMouseDown;
            this.MouseLeftButtonUp += OnMouseUp;
            this.MouseMove += OnMouseMove;
            this.MouseWheel += OnMouseWheel;
            this.MouseLeave += (s, e) => { _isDragging = false; this.ReleaseMouseCapture(); };
        }

        private void OnMouseDown(object sender, MouseButtonEventArgs e)
        {
            _isDragging = true;
            _lastMousePosition = e.GetPosition(this);
            this.CaptureMouse();
        }

        private void OnMouseUp(object sender, MouseButtonEventArgs e)
        {
            _isDragging = false;
            this.ReleaseMouseCapture();
        }

        private void OnMouseMove(object sender, MouseEventArgs e)
        {
            if (!_isDragging) return;

            Point currentPosition = e.GetPosition(this);
            Vector delta = currentPosition - _lastMousePosition;

            _cameraTheta -= delta.X * 0.01;
            _cameraPhi += delta.Y * 0.01;

            if (_cameraPhi < 0.01) _cameraPhi = 0.01;
            if (_cameraPhi > Math.PI / 2 - 0.01) _cameraPhi = Math.PI / 2 - 0.01;

            UpdateCameraPosition();
            _lastMousePosition = currentPosition;
        }

        private void OnMouseWheel(object sender, MouseWheelEventArgs e)
        {
            _cameraRadius -= e.Delta * 0.05;
            if (_cameraRadius < 20) _cameraRadius = 20;
            if (_cameraRadius > 500) _cameraRadius = 500;
            UpdateCameraPosition();
        }

        private void UpdateCameraPosition()
        {
            double x = _cameraRadius * Math.Sin(_cameraPhi) * Math.Cos(_cameraTheta);
            double y = _cameraRadius * Math.Sin(_cameraPhi) * Math.Sin(_cameraTheta);
            double z = _cameraRadius * Math.Cos(_cameraPhi);

            _camera.Position = new Point3D(x, y, z);
            _camera.LookDirection = new Vector3D(-x, -y, -z);
        }

        // ======================= SCENE UPDATERS =======================

        public void UpdateScene_BanKinhRe(double rOut, double u, double carWidth)
        {
            ClearScene();
            double rIn = rOut - u;
            if (rIn <= 0.1) rIn = 0.1;

            double halfW = 10.0; 
            double rKerb = rIn;
            double cx = halfW + rKerb;
            double cy = -halfW - rKerb; 

            Color roadColor = Color.FromRgb(55, 60, 65);
            Color grassColor = Color.FromRgb(40, 90, 45); 
            Color sidewalkColor = Color.FromRgb(180, 180, 180); 
            Color carColor = Color.FromRgb(241, 196, 15); 
            Color pathColor = Color.FromArgb(130, 52, 152, 219); 
            Color medianColor = Color.FromRgb(120, 130, 120);

            _mainGroup.Children.Add(CreateBox(-150, -150, -0.6, 300, 300, 0.5, grassColor));

            _mainGroup.Children.Add(CreateBox(-150, -halfW, -0.1, 300, halfW*2, 0.2, roadColor)); 
            _mainGroup.Children.Add(CreateBox(-halfW, -150, -0.1, halfW*2, 300, 0.2, roadColor)); 

            double inf = 150;
            // NE
            _mainGroup.Children.Add(CreateCurvedRoad(halfW+rKerb, halfW+rKerb, rKerb, rKerb, inf, inf, 0.3, sidewalkColor, 180, 270));
            _mainGroup.Children.Add(CreateBox(halfW, halfW+rKerb, 0.1, rKerb, inf, 0.3, sidewalkColor));
            _mainGroup.Children.Add(CreateBox(halfW+rKerb, halfW, 0.1, inf, rKerb, 0.3, sidewalkColor));
            _mainGroup.Children.Add(CreateBox(halfW+rKerb, halfW+rKerb, 0.1, inf, inf, 0.3, sidewalkColor));
            // NW
            _mainGroup.Children.Add(CreateCurvedRoad(-halfW-rKerb, halfW+rKerb, rKerb, rKerb, inf, inf, 0.3, sidewalkColor, 270, 360));
            _mainGroup.Children.Add(CreateBox(-halfW-rKerb, halfW+rKerb, 0.1, rKerb, inf, 0.3, sidewalkColor));
            _mainGroup.Children.Add(CreateBox(-inf, halfW, 0.1, inf-halfW-rKerb, rKerb, 0.3, sidewalkColor));
            _mainGroup.Children.Add(CreateBox(-inf, halfW+rKerb, 0.1, inf-halfW-rKerb, inf, 0.3, sidewalkColor));
            // SW
            _mainGroup.Children.Add(CreateCurvedRoad(-halfW-rKerb, -halfW-rKerb, rKerb, rKerb, inf, inf, 0.3, sidewalkColor, 0, 90));
            _mainGroup.Children.Add(CreateBox(-halfW-rKerb, -inf, 0.1, rKerb, inf-halfW-rKerb, 0.3, sidewalkColor));
            _mainGroup.Children.Add(CreateBox(-inf, -halfW-rKerb, 0.1, inf-halfW-rKerb, rKerb, 0.3, sidewalkColor));
            _mainGroup.Children.Add(CreateBox(-inf, -inf, 0.1, inf-halfW-rKerb, inf-halfW-rKerb, 0.3, sidewalkColor));
            // SE
            _mainGroup.Children.Add(CreateCurvedRoad(cx, cy, rKerb, rKerb, inf, inf, 0.3, sidewalkColor, 90, 180));
            _mainGroup.Children.Add(CreateBox(halfW, -inf, 0.1, rKerb, inf-halfW-rKerb, 0.3, sidewalkColor));
            _mainGroup.Children.Add(CreateBox(cx, cy, 0.1, inf, rKerb, 0.3, sidewalkColor));
            _mainGroup.Children.Add(CreateBox(cx, -inf, 0.1, inf, inf-halfW-rKerb, 0.3, sidewalkColor));

            // Medians
            double medW = 2.0;
            _mainGroup.Children.Add(CreateBox(-medW/2, halfW, 0.1, medW, inf, 0.3, medianColor));
            _mainGroup.Children.Add(CreateBox(-medW/2, -inf, 0.1, medW, inf-halfW, 0.3, medianColor));
            _mainGroup.Children.Add(CreateBox(halfW, -medW/2, 0.1, inf, medW, 0.3, medianColor));
            _mainGroup.Children.Add(CreateBox(-inf, -medW/2, 0.1, inf-halfW, medW, 0.3, medianColor));

            // SWEPT PATH
            _mainGroup.Children.Add(CreateCurvedRoad(cx, cy, rIn, rIn, rOut, rOut, 0.4, pathColor, 90, 180));
            
            _mainGroup.Children.Add(CreateBox(cx, cy+rIn, 0.3, inf, rOut-rIn, 0.1, pathColor)); 
            _mainGroup.Children.Add(CreateBox(cx-rOut, -inf+50, 0.3, rOut-rIn, inf-50-halfW-rKerb, 0.1, pathColor)); 

            // CARS 
            double carL = 8.0; 
            double startX = cx - (rOut+rIn)/2;
            double startY = cy - 25;
            _mainGroup.Children.Add(CreateVehicleModel(startX, startY, 0.4, carWidth, carL, carColor, 90)); 

            double endX = cx + 25;
            double endY = cy + (rOut+rIn)/2;
            _mainGroup.Children.Add(CreateVehicleModel(endX, endY, 0.4, carWidth, carL, carColor, 0)); 

            double midA = 135 * Math.PI / 180.0;
            double midX = cx + (rOut+rIn)/2 * Math.Cos(midA);
            double midY = cy + (rOut+rIn)/2 * Math.Sin(midA);
            _mainGroup.Children.Add(CreateVehicleModel(midX, midY, 0.4, carWidth, carL, carColor, -45));

            _cameraRadius = Math.Max(rOut * 2.5, 90);
            _cameraPhi = Math.PI / 4;
            UpdateCameraPosition();
        }


        public void UpdateScene_MoRong(double r_xe, double b1, double e_req)
        {
            ClearScene();
            double rIn = r_xe;
            double cx = rIn + b1 + e_req + 20;
            double cy = rIn + b1 + e_req + 20;

            Color roadColor = Color.FromRgb(60, 65, 70);
            Color grassColor = Color.FromRgb(40, 50, 45);
            Color carColor = Color.FromRgb(185, 41, 41); // Red car

            // The curved road widens from b1 to b1+e_req and back to b1
            // We can draw it as a single sweep changing outer r, but actually the transition is on the straights too.
            // Simplified: The full curve has widened lane, straight has normal lane.
            double rOutNormal = rIn + b1;
            double rOutWidened = rIn + b1 + e_req;

            _mainGroup.Children.Add(CreateCurvedRoad(cx, cy, rIn, rIn, rOutNormal, rOutWidened, 1.0, roadColor, 180, 225));
            _mainGroup.Children.Add(CreateCurvedRoad(cx, cy, rIn, rIn, rOutWidened, rOutNormal, 1.0, roadColor, 225, 270));

            // Straight Roads with normal width
            _mainGroup.Children.Add(CreateBox(cx - rOutNormal, cy, 0, b1, 60, 1.0, roadColor));
            _mainGroup.Children.Add(CreateBox(cx, cy - rOutNormal, 0, 60, b1, 1.0, roadColor));

            // Grass
            _mainGroup.Children.Add(CreateBox(cx - rOutWidened - 50, cy - rOutWidened - 50, -0.5, rOutWidened * 2 + 100, rOutWidened * 2 + 100, 0.5, grassColor));

            // Vehicle drifting the wide curve! Position it at 45 deg center.
            double carWidth = 2.0;
            double carLength = 12.0; // Long truck!
            double angle = 225 * Math.PI / 180.0;
            double rCenterCar = rIn + (b1 + e_req) / 2;
            double truckX = cx + rCenterCar * Math.Cos(angle) - carLength / 2;
            double truckY = cy + rCenterCar * Math.Sin(angle) - carWidth / 2;

            // Truck rendering
            _mainGroup.Children.Add(CreateVehicleGroup(truckX, truckY, 1.0, carLength, carWidth, carColor));

            _cameraRadius = Math.Max(rOutWidened * 2.5, 120);
            UpdateCameraPosition();
        }

        private void ClearScene()
        {
            while (_mainGroup.Children.Count > 3)
            {
                _mainGroup.Children.RemoveAt(_mainGroup.Children.Count - 1);
            }
        }

        public void UpdateScene(double radius, double carWidth, double carLength)
        {
            UpdateScene_BanKinhRe(radius, carWidth, carLength);
        }

        public void UpdateScene_VongXuyen(double d_dt, double r_xe, double e1)
        {
            ClearScene();
            double rIsland = d_dt / 2.0;
            double rOut = r_xe + e1; // Approximation of outer edge
            double cx = 0;
            double cy = 0;

            Color roadColor = Color.FromRgb(60, 65, 70);
            Color grassColor = Color.FromRgb(40, 50, 45);
            Color islandColor = Color.FromRgb(50, 100, 50); // Green island

            // Central Island
            _mainGroup.Children.Add(CreateCylinder(cx, cy, 0, rIsland, 2.0, islandColor));

            // Circulating Road (Annulus)
            _mainGroup.Children.Add(CreateCurvedRoad(cx, cy, rIsland, rIsland, rOut, rOut, 1.0, roadColor, 0, 360));

            // Leg width
            double legWidth = 15.0; // Approximation

            // 4 Intersecting Legs
            _mainGroup.Children.Add(CreateBox(cx - legWidth / 2, cy + rOut, 0, legWidth, 50, 1.0, roadColor)); // Top
            _mainGroup.Children.Add(CreateBox(cx - legWidth / 2, cy - rOut - 50, 0, legWidth, 50, 1.0, roadColor)); // Bottom
            _mainGroup.Children.Add(CreateBox(cx + rOut, cy - legWidth / 2, 0, 50, legWidth, 1.0, roadColor)); // Right
            _mainGroup.Children.Add(CreateBox(cx - rOut - 50, cy - legWidth / 2, 0, 50, legWidth, 1.0, roadColor)); // Left

            // Ground
            _mainGroup.Children.Add(CreateBox(-100, -100, -0.5, 200, 200, 0.5, grassColor));

            // Vehicles circulating
            double carW = 2.0;
            double carL = 4.5;
            Color carC1 = Color.FromRgb(200, 50, 50);
            Color carC2 = Color.FromRgb(50, 50, 200);

            // Vehicle 1 at 45 deg
            double a1 = 45 * Math.PI / 180;
            double v1x = cx + r_xe * Math.Cos(a1) - carW / 2; // rough pos
            double v1y = cy + r_xe * Math.Sin(a1) - carL / 2;
            _mainGroup.Children.Add(CreateVehicleGroup(v1x, v1y, 1.0, carL, carW, carC1));

            // Vehicle 2 at 135 deg
            double a2 = 135 * Math.PI / 180;
            double v2x = cx + r_xe * Math.Cos(a2) - carL / 2; // rotated 90 approx
            double v2y = cy + r_xe * Math.Sin(a2) - carW / 2;
            _mainGroup.Children.Add(CreateVehicleGroup(v2x, v2y, 1.0, carW, carL, carC2));

            _cameraRadius = Math.Max(rOut * 3.5, 120);
            UpdateCameraPosition();
        }


        public void UpdateScene_TamNhin(double sa, double sb)
        {
            ClearScene();
            double roadWidth = 20.0;
            double cx = 0;
            double cy = 0;

            Color roadColor = Color.FromRgb(60, 65, 70);
            Color grassColor = Color.FromRgb(40, 50, 45);
            Color triangleColor = Color.FromArgb(150, 255, 50, 50); // Semi-transparent Red

            // 4 Intersecting Legs
            double ext = Math.Max(sa, sb) + 50;
            _mainGroup.Children.Add(CreateBox(cx - roadWidth / 2, cy, 0, roadWidth, ext, 1.0, roadColor)); // Top
            _mainGroup.Children.Add(CreateBox(cx - roadWidth / 2, cy - ext, 0, roadWidth, ext, 1.0, roadColor)); // Bottom
            _mainGroup.Children.Add(CreateBox(cx, cy - roadWidth / 2, 0, ext, roadWidth, 1.0, roadColor)); // Right
            _mainGroup.Children.Add(CreateBox(cx - ext, cy - roadWidth / 2, 0, ext, roadWidth, 1.0, roadColor)); // Left

            // Ground
            _mainGroup.Children.Add(CreateBox(-ext - 20, -ext - 20, -0.5, ext * 2 + 40, ext * 2 + 40, 0.5, grassColor));

            // Sight Triangle
            double aX = cx + roadWidth / 4;
            double aY = cy - sa;
            double bX = cx + sb;
            double bY = cy - roadWidth / 4;

            _mainGroup.Children.Add(CreateTriangle(aX, aY, bX, bY, cx, cy, 1.5, triangleColor));

            // Vehicles
            Color carC = Color.FromRgb(41, 128, 185);
            _mainGroup.Children.Add(CreateVehicleGroup(aX - 1.0, aY - 4.5, 1.0, 2.0, 4.5, carC)); // Vertical
            _mainGroup.Children.Add(CreateVehicleGroup(bX, bY - 1.0, 1.0, 4.5, 2.0, carC)); // Horizontal

            _cameraRadius = Math.Max(ext * 2.0, 150);
            UpdateCameraPosition();
        }

        public void UpdateScene_TangGiamToc(bool isTangToc, double ln, double lv)
        {
            ClearScene();
            double mainW = 20.0;
            double sideW = 15.0; // Vuốt / làn thêm

            Color roadColor = Color.FromRgb(60, 65, 70);
            Color grassColor = Color.FromRgb(40, 50, 45);
            Color carC = Color.FromRgb(41, 128, 185);

            double totalL = ln + lv + 100;

            // Main Road
            _mainGroup.Children.Add(CreateBox(0, 0, 0, totalL, mainW, 1.0, roadColor));

            // Ground
            _mainGroup.Children.Add(CreateBox(-20, -20, -0.5, totalL + 40, mainW + sideW + 40, 0.5, grassColor));

            double startX = 50;
            if (isTangToc)
            {
                // Parallel lane first (Lv) then merges (Ln)
                _mainGroup.Children.Add(CreateBox(startX, mainW, 0, lv, sideW, 1.0, roadColor));
                _mainGroup.Children.Add(CreateCustomQuad(startX + lv, mainW, startX + lv, mainW + sideW, startX + lv + ln, mainW, startX + lv + ln, mainW, 1.0, roadColor));

                // Car merging
                _mainGroup.Children.Add(CreateVehicleGroup(startX + lv - 10, mainW + sideW / 2 - 1.0, 1.0, 4.5, 2.0, carC));
            }
            else
            {
                // Diverges (Ln) then parallel (Lv)
                _mainGroup.Children.Add(CreateCustomQuad(startX, mainW, startX, mainW, startX + ln, mainW, startX + ln, mainW + sideW, 1.0, roadColor));
                _mainGroup.Children.Add(CreateBox(startX + ln, mainW, 0, lv, sideW, 1.0, roadColor));

                // Car diverging
                _mainGroup.Children.Add(CreateVehicleGroup(startX + ln + 10, mainW + sideW / 2 - 1.0, 1.0, 4.5, 2.0, carC));
            }

            _cameraRadius = Math.Max(totalL * 1.2, 100);
            UpdateCameraPosition();
        }

        public void UpdateScene_LanRe(bool isLeftTurn, double ln, double lv, double lx)
        {
            ClearScene();
            double mainW = 20.0;
            double sideW = 12.0;

            Color roadColor = Color.FromRgb(60, 65, 70);
            Color grassColor = Color.FromRgb(40, 50, 45);
            Color carC = Color.FromRgb(41, 128, 185);
            Color redCarC = Color.FromRgb(185, 41, 41);

            double totalL = ln + lv + lx + 100;

            // Main Road
            _mainGroup.Children.Add(CreateBox(0, 0, 0, totalL, mainW, 1.0, roadColor));

            // Ground
            _mainGroup.Children.Add(CreateBox(-20, -20, -0.5, totalL + 40, mainW + sideW + 40, 0.5, grassColor));

            double startX = 50;

            if (isLeftTurn)
            {
                // Left turn lane (above main lane)
                _mainGroup.Children.Add(CreateCustomQuad(startX, mainW, startX, mainW, startX + ln, mainW, startX + ln, mainW + sideW, 1.0, roadColor));
                _mainGroup.Children.Add(CreateBox(startX + ln, mainW, 0, lv + lx, sideW, 1.0, roadColor));

                // Cars queuing in lx
                _mainGroup.Children.Add(CreateVehicleGroup(startX + ln + lv + 5, mainW + sideW / 2 - 1.0, 1.0, 4.5, 2.0, redCarC));
                _mainGroup.Children.Add(CreateVehicleGroup(startX + ln + lv + 12, mainW + sideW / 2 - 1.0, 1.0, 4.5, 2.0, redCarC));
            }
            else
            {
                // Right turn lane (below main lane)
                _mainGroup.Children.Add(CreateCustomQuad(startX, 0, startX, 0, startX + ln, -sideW, startX + ln, 0, 1.0, roadColor));
                _mainGroup.Children.Add(CreateBox(startX + ln, -sideW, 0, lv + lx, sideW, 1.0, roadColor));

                // Cars queuing in lx
                _mainGroup.Children.Add(CreateVehicleGroup(startX + ln + lv + 5, -sideW / 2 - 1.0, 1.0, 4.5, 2.0, carC));
                _mainGroup.Children.Add(CreateVehicleGroup(startX + ln + lv + 12, -sideW / 2 - 1.0, 1.0, 4.5, 2.0, carC));
            }

            _cameraRadius = Math.Max(totalL * 1.2, 100);
            UpdateCameraPosition();
        }

        private Model3DGroup CreateVehicleGroup(double x, double y, double z, double w, double l, Color color)
        {
            // Fallback for older methods that use absolute coordinates
            return CreateVehicleModel(x + w/2, y + l/2, z, w, l, color, (w < l) ? 90 : 0);
        }

        private Model3DGroup CreateVehicleModel(double cx, double cy, double z, double carW, double carL, Color color, double headingDeg)
        {
            Model3DGroup carGroup = new Model3DGroup();
            double hw = carW / 2;
            double hl = carL / 2;

            carGroup.Children.Add(CreateBox(-hl, -hw, z + 0.5, carL, carW, 2.0, color));
            
            double cabL = carL * 0.4;
            double cabW = carW * 0.8;
            carGroup.Children.Add(CreateBox(hl - cabL, -cabW/2, z + 2.5, cabL, cabW, 1.2, Color.FromRgb(40, 40, 50))); // Cab at front

            Color wheelC = Color.FromRgb(15, 15, 15);
            double wl = 1.0; 
            double ww = 0.5;
            double wx = hl * 0.6;
            double wy = hw + 0.1;
            carGroup.Children.Add(CreateBox(-wx, wy, z, wl, ww, 1.0, wheelC)); 
            carGroup.Children.Add(CreateBox(-wx, -wy-ww, z, wl, ww, 1.0, wheelC)); 
            carGroup.Children.Add(CreateBox(wx - wl, wy, z, wl, ww, 1.0, wheelC)); 
            carGroup.Children.Add(CreateBox(wx - wl, -wy-ww, z, wl, ww, 1.0, wheelC)); 

            Transform3DGroup tfg = new Transform3DGroup();
            tfg.Children.Add(new RotateTransform3D(new AxisAngleRotation3D(new Vector3D(0, 0, 1), headingDeg)));
            tfg.Children.Add(new TranslateTransform3D(cx, cy, 0));
            carGroup.Transform = tfg;

            return carGroup;
        }

        private Model3DGroup CreateVehicle(double x, double y, double z, double w, double l, Color color)
        {
            return CreateVehicleGroup(x, y, z, w, l, color);
        }

        private GeometryModel3D CreateCustomQuad(double x1, double y1, double x2, double y2, double x3, double y3, double x4, double y4, double height, Color c)
        {
            MeshGeometry3D mesh = new MeshGeometry3D();

            mesh.Positions.Add(new Point3D(x1, y1, 0));
            mesh.Positions.Add(new Point3D(x2, y2, 0));
            mesh.Positions.Add(new Point3D(x3, y3, 0));
            mesh.Positions.Add(new Point3D(x4, y4, 0));

            mesh.Positions.Add(new Point3D(x1, y1, height));
            mesh.Positions.Add(new Point3D(x2, y2, height));
            mesh.Positions.Add(new Point3D(x3, y3, height));
            mesh.Positions.Add(new Point3D(x4, y4, height));

            AddQuad(mesh, 0, 1, 2, 3); // Bottom
            AddQuad(mesh, 4, 5, 6, 7); // Top

            Material mat = new DiffuseMaterial(new SolidColorBrush(c));
            return new GeometryModel3D(mesh, mat);
        }

        private GeometryModel3D CreateBox(double x, double y, double z, double w, double l, double h, Color c)
        {
            MeshGeometry3D mesh = new MeshGeometry3D();

            mesh.Positions.Add(new Point3D(x, y, z));
            mesh.Positions.Add(new Point3D(x + w, y, z));
            mesh.Positions.Add(new Point3D(x + w, y + l, z));
            mesh.Positions.Add(new Point3D(x, y + l, z));
            mesh.Positions.Add(new Point3D(x, y, z + h));
            mesh.Positions.Add(new Point3D(x + w, y, z + h));
            mesh.Positions.Add(new Point3D(x + w, y + l, z + h));
            mesh.Positions.Add(new Point3D(x, y + l, z + h));

            AddQuad(mesh, 0, 1, 5, 4);
            AddQuad(mesh, 1, 2, 6, 5);
            AddQuad(mesh, 2, 3, 7, 6);
            AddQuad(mesh, 3, 0, 4, 7);
            AddQuad(mesh, 4, 5, 6, 7);
            AddQuad(mesh, 3, 2, 1, 0);

            Material mat = new DiffuseMaterial(new SolidColorBrush(c));
            return new GeometryModel3D(mesh, mat);
        }

        private GeometryModel3D CreateTriangle(double x1, double y1, double x2, double y2, double x3, double y3, double z, Color c)
        {
            MeshGeometry3D mesh = new MeshGeometry3D();

            mesh.Positions.Add(new Point3D(x1, y1, z));
            mesh.Positions.Add(new Point3D(x2, y2, z));
            mesh.Positions.Add(new Point3D(x3, y3, z));

            // Double sided
            mesh.TriangleIndices.Add(0);
            mesh.TriangleIndices.Add(1);
            mesh.TriangleIndices.Add(2);

            mesh.TriangleIndices.Add(0);
            mesh.TriangleIndices.Add(2);
            mesh.TriangleIndices.Add(1);

            Material diffuseMat = new DiffuseMaterial(new SolidColorBrush(c));

            // To support transparency properly if color has alpha
            MaterialGroup matGroup = new MaterialGroup();
            matGroup.Children.Add(diffuseMat);

            return new GeometryModel3D(mesh, matGroup);
        }

        private GeometryModel3D CreateCylinder(double cx, double cy, double z, double radius, double height, Color c)
        {
            MeshGeometry3D mesh = new MeshGeometry3D();
            int segments = 40;

            mesh.Positions.Add(new Point3D(cx, cy, z + height)); // Top center = 0
            mesh.Positions.Add(new Point3D(cx, cy, z));         // Bottom center = 1

            for (int i = 0; i <= segments; i++)
            {
                double angle = (Math.PI * 2.0 * i) / segments;
                double cos = Math.Cos(angle);
                double sin = Math.Sin(angle);

                mesh.Positions.Add(new Point3D(cx + radius * cos, cy + radius * sin, z + height)); // Top edge
                mesh.Positions.Add(new Point3D(cx + radius * cos, cy + radius * sin, z));          // Bottom edge 
            }

            // Indices
            for (int i = 0; i < segments; i++)
            {
                int top1 = 2 + i * 2;
                int bot1 = 3 + i * 2;
                int top2 = 2 + (i + 1) * 2;
                int bot2 = 3 + (i + 1) * 2;

                // Top triangle (CCW)
                mesh.TriangleIndices.Add(0);
                mesh.TriangleIndices.Add(top1);
                mesh.TriangleIndices.Add(top2);

                // Bottom triangle (CW from top)
                mesh.TriangleIndices.Add(1);
                mesh.TriangleIndices.Add(bot2);
                mesh.TriangleIndices.Add(bot1);

                // Side quads
                AddQuad(mesh, bot1, bot2, top2, top1);
            }

            Material mat = new DiffuseMaterial(new SolidColorBrush(c));
            return new GeometryModel3D(mesh, mat);
        }

        private GeometryModel3D CreateCurvedRoad(double cx, double cy, double rInStart, double rInEnd, double rOutStart, double rOutEnd, double thickness, Color c, double startAngleDeg, double sweepAngleDeg)
        {
            MeshGeometry3D mesh = new MeshGeometry3D();
            int segments = 40;

            double startRad = startAngleDeg * Math.PI / 180.0;
            // Target angle (absolute)
            double endRad = sweepAngleDeg * Math.PI / 180.0;
            // To make sweepAngle work like target angle or sweep...? 
            // In Wpf3DViewer old code it was target angle but named sweep. 
            // Let's assume start = 180, end = 270 is absolute.
            // Oh wait, my old code: angle = startRad + sweepRad * ... No, my old code passed 270 as sweepAngle! That means 180 + 270 = 450 degrees! 
            // Wait, let's fix that. If I pass start=180, sweep=90, then angle = 180 + 90*(i/segments) = 270.
            double sweepRad = sweepAngleDeg * Math.PI / 180.0 - startRad; // If sweepAngle == targetAngle. Wait!
            // Let's explicitly define startAngle and endAngle!
            double endAngleDeg = sweepAngleDeg; // renamed parameter usage
            sweepRad = (endAngleDeg - startAngleDeg) * Math.PI / 180.0;

            for (int i = 0; i <= segments; i++)
            {
                double t = (double)i / segments;
                double angle = startRad + sweepRad * t;
                double cos = Math.Cos(angle);
                double sin = Math.Sin(angle);

                double currentRIn = rInStart + (rInEnd - rInStart) * t;
                double currentROut = rOutStart + (rOutEnd - rOutStart) * t;

                mesh.Positions.Add(new Point3D(cx + currentRIn * cos, cy + currentRIn * sin, 0));
                mesh.Positions.Add(new Point3D(cx + currentROut * cos, cy + currentROut * sin, 0));
                mesh.Positions.Add(new Point3D(cx + currentRIn * cos, cy + currentRIn * sin, thickness));
                mesh.Positions.Add(new Point3D(cx + currentROut * cos, cy + currentROut * sin, thickness));
            }

            for (int i = 0; i < segments; i++)
            {
                int b1 = i * 4;
                int b2 = (i + 1) * 4;

                AddQuad(mesh, b1 + 2, b1 + 3, b2 + 3, b2 + 2); // Top
                AddQuad(mesh, b2 + 1, b2 + 0, b1 + 0, b1 + 1); // Bottom
                AddQuad(mesh, b1 + 0, b2 + 0, b2 + 2, b1 + 2); // Inner
                AddQuad(mesh, b1 + 3, b2 + 3, b2 + 1, b1 + 1); // Outer
            }

            Material mat = new DiffuseMaterial(new SolidColorBrush(c));
            return new GeometryModel3D(mesh, mat);
        }

        private void AddQuad(MeshGeometry3D mesh, int p1, int p2, int p3, int p4)
        {
            mesh.TriangleIndices.Add(p1);
            mesh.TriangleIndices.Add(p2);
            mesh.TriangleIndices.Add(p3);

            mesh.TriangleIndices.Add(p1);
            mesh.TriangleIndices.Add(p3);
            mesh.TriangleIndices.Add(p4);
        }
    }
}
