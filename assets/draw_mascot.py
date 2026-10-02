"""用法：python assets/draw_mascot.py assets

原创 Q 版小人：绿色长发、齐刘海、金色半睁眼、抱着一根黄瓜。全部用矢量路径画，再加白色贴纸描边。"""
import sys, os
import numpy as np
from PySide6.QtCore import Qt, QPointF, QRectF
from PySide6.QtGui import QGuiApplication, QImage, QPainter, QColor, QPainterPath, QPen, QLinearGradient, QRadialGradient
from scipy import ndimage

out = sys.argv[1]
app = QGuiApplication(sys.argv)
W, H = 640, 960
img = QImage(W, H, QImage.Format_ARGB32_Premultiplied)
img.fill(Qt.transparent)
p = QPainter(img)
p.setRenderHints(QPainter.Antialiasing)

LINE = QColor(74, 92, 70)            # 统一的深绿灰线条
HAIR = QColor(176, 204, 166)
HAIR_D = QColor(138, 172, 130)
HAIR_L = QColor(214, 232, 204)
SKIN = QColor(255, 236, 224)
SKIN_D = QColor(244, 206, 192)
DRESS = QColor(124, 158, 122)
DRESS_D = QColor(98, 130, 98)
WHITE = QColor(252, 251, 245)
GOLD = QColor(232, 186, 82)
GOLD_D = QColor(176, 128, 44)


def path(points, close=True):
    pp = QPainterPath(QPointF(*points[0]))
    for seg in points[1:]:
        if len(seg) == 2:
            pp.lineTo(*seg)
        elif len(seg) == 4:
            pp.quadTo(seg[0], seg[1], seg[2], seg[3])
        else:
            pp.cubicTo(*seg)
    if close:
        pp.closeSubpath()
    return pp


def fill(pp, color, line=True, width=5):
    p.setPen(QPen(LINE, width, Qt.SolidLine, Qt.RoundCap, Qt.RoundJoin) if line else Qt.NoPen)
    p.setBrush(color)
    p.drawPath(pp)


cx = 320

# ---- 后发（长到腰下）
back = path([(cx - 190, 300), (cx - 230, 520, cx - 240, 700, cx - 200, 760), (cx - 150, 790, cx - 120, 700),
             (cx + 120, 700), (cx + 150, 790, cx + 200, 760), (cx + 240, 700, cx + 230, 520, cx + 190, 300),
             (cx + 120, 110, cx - 120, 110, cx - 190, 300)])
g = QLinearGradient(0, 150, 0, 780)
g.setColorAt(0, HAIR)
g.setColorAt(1, HAIR_D)
fill(back, g)
# 后发的发丝线
p.setPen(QPen(QColor(110, 140, 104, 150), 3, Qt.SolidLine, Qt.RoundCap))
p.setBrush(Qt.NoBrush)
for side in (-1, 1):
    for k, (x0, x1) in enumerate(((175, 205), (150, 172), (128, 140))):
        p.drawPath(path([(cx + side * x0, 330), (cx + side * (x1 + 18), 520, cx + side * x1, 700 - k * 30)], close=False))

# ---- 腿和鞋
for dx in (-48, 48):
    fill(path([(cx + dx - 26, 770), (cx + dx - 24, 860), (cx + dx + 24, 860), (cx + dx + 26, 770)]), SKIN)
    fill(path([(cx + dx - 28, 830), (cx + dx - 28, 862), (cx + dx + 28, 862), (cx + dx + 28, 830)]), WHITE)
    fill(path([(cx + dx - 34, 860), (cx + dx - 36, 900, cx + dx + 36, 900, cx + dx + 34, 860)]), QColor(70, 72, 82))

# ---- 裙子
dress = path([(cx - 100, 560), (cx - 170, 790), (cx - 60, 812, cx + 60, 812, cx + 170, 790), (cx + 100, 560)])
g = QLinearGradient(0, 560, 0, 810)
g.setColorAt(0, DRESS)
g.setColorAt(1, DRESS_D)
fill(dress, g)
# 裙摆白边
fill(path([(cx - 166, 778), (cx - 60, 800, cx + 60, 800, cx + 166, 778), (cx + 172, 796),
           (cx + 60, 822, cx - 60, 822, cx - 172, 796)]), WHITE, width=4)
# 上身
fill(path([(cx - 100, 560), (cx - 92, 470, cx + 92, 470, cx + 100, 560)]), DRESS)
# 白领子
fill(path([(cx - 70, 478), (cx, 540), (cx + 70, 478), (cx + 40, 466), (cx, 500), (cx - 40, 466)]), WHITE, width=4)
# 领结（金色）
fill(path([(cx, 520), (cx - 34, 500), (cx - 34, 546)]), GOLD, width=4)
fill(path([(cx, 520), (cx + 34, 500), (cx + 34, 546)]), GOLD, width=4)
p.setBrush(GOLD_D)
p.drawEllipse(QPointF(cx, 522), 10, 10)

# ---- 黄瓜（抱在胸前，斜着）
p.save()
p.translate(cx + 6, 640)
p.rotate(-28)
cuc = QPainterPath()
cuc.addRoundedRect(QRectF(-150, -30, 300, 60), 30, 30)
g = QLinearGradient(0, -30, 0, 30)
g.setColorAt(0, QColor(126, 186, 92))
g.setColorAt(0.5, QColor(92, 150, 64))
g.setColorAt(1, QColor(62, 112, 46))
fill(cuc, g)
p.setPen(Qt.NoPen)
p.setBrush(QColor(196, 230, 150, 200))
for x in range(-120, 130, 34):                 # 小刺点
    p.drawEllipse(QPointF(x, -10 + (x // 34 % 2) * 14), 4, 4)
p.setBrush(QColor(236, 214, 96))
p.setPen(QPen(LINE, 3))
p.drawEllipse(QPointF(150, 0), 14, 12)          # 小黄花
p.restore()

# ---- 小手（抱住黄瓜）
for hx, hy in ((cx - 80, 670), (cx + 76, 604)):
    fill(path([(hx - 30, hy), (hx - 30, hy - 34, hx + 30, hy - 34, hx + 30, hy), (hx + 30, hy + 30, hx - 30, hy + 30, hx - 30, hy)]), SKIN, width=4)
# 袖子
fill(path([(cx - 100, 490), (cx - 150, 600), (cx - 112, 660), (cx - 80, 560)]), DRESS, width=5)
fill(path([(cx + 100, 490), (cx + 150, 560), (cx + 100, 600), (cx + 80, 540)]), DRESS, width=5)

# ---- 脸
face = QPainterPath()
face.addEllipse(QRectF(cx - 165, 175, 330, 300))
fill(face, SKIN)
# 腮红
for bx in (cx - 100, cx + 100):
    rg = QRadialGradient(QPointF(bx, 385), 38)
    rg.setColorAt(0, QColor(255, 170, 170, 150))
    rg.setColorAt(1, QColor(255, 170, 170, 0))
    p.setPen(Qt.NoPen)
    p.setBrush(rg)
    p.drawEllipse(QPointF(bx, 385), 38, 26)

# ---- 眼睛：金色、上眼睑压低 = 没什么表情的半睁眼
for ex in (cx - 70, cx + 70):
    eye = QRectF(ex - 40, 290, 80, 96)
    clip = QPainterPath()
    clip.addRect(QRectF(ex - 50, 313, 100, 85))      # 上眼睑压低（眼珠藏到眼睑线下面一点）
    iris = QPainterPath()
    iris.addEllipse(eye)
    p.save()
    p.setClipPath(clip)
    g = QLinearGradient(0, 300, 0, 386)
    g.setColorAt(0, GOLD_D)
    g.setColorAt(0.55, GOLD)
    g.setColorAt(1, QColor(250, 222, 140))
    p.setPen(Qt.NoPen)
    p.setBrush(g)
    p.drawPath(iris)
    p.setBrush(QColor(120, 84, 30))
    p.drawEllipse(QPointF(ex, 352), 18, 24)          # 瞳孔
    p.setBrush(QColor(255, 255, 255, 235))
    p.drawEllipse(QPointF(ex - 14, 336), 9, 9)       # 高光
    p.drawEllipse(QPointF(ex + 14, 368), 5, 5)
    p.restore()
    p.setPen(QPen(LINE, 7, Qt.SolidLine, Qt.RoundCap))
    p.drawLine(QPointF(ex - 46, 318), QPointF(ex + 46, 318))   # 平平的上眼睑
    p.setPen(QPen(LINE, 3, Qt.SolidLine, Qt.RoundCap))
    p.drawArc(QRectF(ex - 36, 300, 72, 92), 200 * 16, 140 * 16)  # 下眼线

# 小嘴：一条短线
p.setPen(QPen(QColor(190, 120, 110), 4, Qt.SolidLine, Qt.RoundCap))
p.drawLine(QPointF(cx - 10, 420), QPointF(cx + 10, 420))

# ---- 齐刘海 + 两侧鬓发
# 平直的齐刘海，只留几处浅浅的发丝缺口
bangs = path([(cx - 175, 300), (cx - 190, 160, cx - 60, 120, cx, 122), (cx + 60, 120, cx + 190, 160, cx + 175, 300),
              (cx + 140, 294), (cx + 118, 284), (cx + 104, 295), (cx + 44, 293), (cx + 30, 283), (cx + 18, 294),
              (cx - 46, 294), (cx - 60, 283), (cx - 72, 294), (cx - 128, 294), (cx - 142, 285), (cx - 152, 296)])
g = QLinearGradient(0, 120, 0, 300)
g.setColorAt(0, HAIR_L)
g.setColorAt(1, HAIR)
fill(bangs, g)
for side in (-1, 1):                                     # 垂在脸两侧的长鬓发
    x0 = cx + side * 160
    fill(path([(x0, 260), (x0 + side * 30, 420, x0 + side * 20, 560, x0 - side * 10, 620),
               (x0 - side * 40, 560, x0 - side * 50, 420, x0 - side * 20, 280)]), HAIR)
# 头发高光
p.setPen(QPen(QColor(255, 255, 255, 170), 6, Qt.SolidLine, Qt.RoundCap))
p.setBrush(Qt.NoBrush)
p.drawArc(QRectF(cx - 120, 150, 240, 120), 40 * 16, 100 * 16)
p.end()

# ---- 白色贴纸描边 + 淡淡的外线
arr = np.frombuffer(img.constBits(), np.uint8).reshape(H, W, 4).copy()   # BGRA (premultiplied)
alpha = arr[:, :, 3] > 20
yy, xx = np.mgrid[-14:15, -14:15]
outline = ndimage.binary_dilation(alpha, structure=(xx ** 2 + yy ** 2) <= 14 ** 2) & ~alpha
outer = ndimage.binary_dilation(outline | alpha, iterations=2) & ~(outline | alpha)
arr[outline] = (245, 250, 250, 255)
arr[outer] = (170, 185, 170, 255)
res = QImage(arr.data, W, H, 4 * W, QImage.Format_ARGB32_Premultiplied).copy()
ys, xs = np.where(arr[:, :, 3] > 0)
res = res.copy(xs.min(), ys.min(), xs.max() - xs.min() + 1, ys.max() - ys.min() + 1)
res = res.scaledToHeight(600, Qt.SmoothTransformation)
res.save(os.path.join(out, "mascot_default.png"))

# 预览
pv = QImage(res.width() * 2 + 40, res.height() + 20, QImage.Format_ARGB32)
pv.fill(QColor(40, 40, 48))
q = QPainter(pv)
q.fillRect(res.width() + 20, 0, res.width() + 20, pv.height(), QColor(235, 120, 90))
q.drawImage(10, 10, res)
q.drawImage(res.width() + 30, 10, res)
q.end()
pv.save(os.path.join(out, "mascot_preview.png"))
print("size:", res.size())
