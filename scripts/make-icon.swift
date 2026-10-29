import AppKit
import Foundation

let size = 1024.0
let image = NSImage(size: NSSize(width: size, height: size))
image.lockFocus()

let blue = NSColor(srgbRed: 0.16, green: 0.46, blue: 0.95, alpha: 1)
let red = NSColor(srgbRed: 0.95, green: 0.27, blue: 0.24, alpha: 1)

// dark body with a blue-to-red gradient edge, matching the Jarvis family
let outer = NSBezierPath(roundedRect: NSRect(x: 8, y: 8, width: size - 16, height: size - 16), xRadius: 226, yRadius: 226)
NSGradient(colors: [blue, red], atLocations: [0, 1], colorSpace: .sRGB)?
  .draw(in: outer, angle: -45)

let body = NSBezierPath(roundedRect: NSRect(x: 26, y: 26, width: size - 52, height: size - 52), xRadius: 210, yRadius: 210)
NSColor(srgbRed: 0.055, green: 0.075, blue: 0.11, alpha: 1).setFill()
body.fill()

// document page
let pageRect = NSRect(x: 286, y: 150, width: 452, height: 668)
let page = NSBezierPath(roundedRect: pageRect, xRadius: 34, yRadius: 34)
NSColor.white.setFill()
page.fill()

// folded corner
let fold = NSBezierPath()
fold.move(to: NSPoint(x: 738 - 118, y: 818))
fold.line(to: NSPoint(x: 738, y: 818 - 118))
fold.line(to: NSPoint(x: 738, y: 818))
fold.close()
NSColor(srgbRed: 0.84, green: 0.87, blue: 0.92, alpha: 1).setFill()
fold.fill()

// blue J mark
let mark = "J" as NSString
let markAttrs: [NSAttributedString.Key: Any] = [
  .font: NSFont.systemFont(ofSize: 380, weight: .bold),
  .foregroundColor: blue
]
let markSize = mark.size(withAttributes: markAttrs)
mark.draw(
  at: NSPoint(x: pageRect.midX - markSize.width / 2, y: 352),
  withAttributes: markAttrs
)

// red rule under the mark
NSColor(srgbRed: 0.95, green: 0.27, blue: 0.24, alpha: 1).setFill()
NSBezierPath(roundedRect: NSRect(x: 362, y: 286, width: 300, height: 26), xRadius: 13, yRadius: 13).fill()

// two body lines
NSColor(srgbRed: 0.80, green: 0.83, blue: 0.88, alpha: 1).setFill()
NSBezierPath(roundedRect: NSRect(x: 362, y: 224, width: 300, height: 22), xRadius: 11, yRadius: 11).fill()
NSBezierPath(roundedRect: NSRect(x: 362, y: 180, width: 196, height: 22), xRadius: 11, yRadius: 11).fill()

image.unlockFocus()

guard
  let tiff = image.tiffRepresentation,
  let rep = NSBitmapImageRep(data: tiff),
  let png = rep.representation(using: .png, properties: [:])
else {
  fputs("failed to encode icon\n", stderr)
  exit(1)
}

let dest = CommandLine.arguments.dropFirst().first ?? "build/icon.png"
try png.write(to: URL(fileURLWithPath: dest))
print(dest)
