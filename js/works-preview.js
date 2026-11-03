import { WebGLRenderer, PerspectiveCamera, Scene, Box3, Vector3, Color } from 'https://traihonhwetw.github.io/ltjh-open-brush-gallery/three.module.js';
import { OrbitControls } from 'https://traihonhwetw.github.io/ltjh-open-brush-gallery/OrbitControls.js';
import { GLTFLoader } from 'https://traihonhwetw.github.io/ltjh-open-brush-gallery/GLTFLoader.js';
import { GLTFGoogleTiltBrushMaterialExtension } from 'https://traihonhwetw.github.io/ltjh-open-brush-gallery/three-icosa.module.js';

const base = 'https://traihonhwetw.github.io/ltjh-open-brush-gallery/';
const requested = new URLSearchParams(location.search).get('model');
const model = ['student-202.glb', 'student-203.glb'].includes(requested) ? requested : 'student-202.glb';
const status = document.getElementById('status');
try {
  const renderer = new WebGLRenderer({ antialias: true });
  renderer.setPixelRatio(Math.min(devicePixelRatio, 1.5));
  renderer.setSize(innerWidth, innerHeight);
  document.body.appendChild(renderer.domElement);
  const scene = new Scene();
  scene.background = new Color('#111410');
  const camera = new PerspectiveCamera(40, innerWidth / innerHeight, 0.01, 10000);
  const controls = new OrbitControls(camera, renderer.domElement);
  controls.autoRotate = false;
  controls.enableZoom = false;
  controls.enablePan = false;
  controls.enableDamping = true;
  const loader = new GLTFLoader();
  loader.register(parser => new GLTFGoogleTiltBrushMaterialExtension(parser, base + 'brushes/'));
  loader.load(base + model, data => {
    const artwork = data.scene;
    artwork.updateMatrixWorld(true);
    const box = new Box3();
    // Floating particle brushes should not push the subject out of the preview.
    artwork.traverse(node => {
      if (node.isMesh && !/brush_(Embers|Snow|Comet|Stars|Streamers)_/i.test(node.name)) {
        box.union(new Box3().setFromObject(node));
      }
    });
    if (box.isEmpty()) box.setFromObject(artwork);
    const center = box.getCenter(new Vector3());
    const size = box.getSize(new Vector3());
    artwork.position.sub(center);
    scene.add(artwork);
    const vertical = Math.max(size.y, size.x / camera.aspect);
    const distance = vertical / (2 * Math.tan(camera.fov * Math.PI / 360)) * 1.25 + size.z / 2;
    camera.position.set(distance * 0.28, distance * 0.12, distance);
    camera.near = Math.max(distance / 1000, 0.001);
    camera.far = distance * 100;
    camera.updateProjectionMatrix();
    controls.target.set(0, 0, 0);
    controls.update();
    status.hidden = true;
  }, undefined, () => { status.textContent = '預覽暫時無法載入，請使用右側的「開啟 3D 作品」。'; });
  addEventListener('resize', () => {
    camera.aspect = innerWidth / innerHeight;
    camera.updateProjectionMatrix();
    renderer.setSize(innerWidth, innerHeight);
  });
  renderer.setAnimationLoop(() => { controls.update(); renderer.render(scene, camera); });
} catch {
  status.textContent = '此瀏覽器無法顯示 3D 預覽，請使用「開啟 3D 作品」。';
}
