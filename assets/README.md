# Assets

This folder contains icons and images for the RST Word Add-in.

## Required Icons

The following icons are required for the add-in to display properly:

| File | Size | Usage |
|------|------|-------|
| `icon-16.png` | 16x16 | Small ribbon icon |
| `icon-32.png` | 32x32 | Standard ribbon icon |
| `icon-80.png` | 80x80 | Large icon, add-in catalog |

## Icon Guidelines

- Use PNG format with transparency
- Design should be clear and recognizable at small sizes
- Recommended: Use a document/text icon with "RST" or similar indicator
- Color: Match Office ribbon style (use Fluent UI colors)

## Generating Placeholder Icons

You can generate simple placeholder icons using ImageMagick:

```bash
# Install ImageMagick if not present
# sudo apt install imagemagick

# Generate placeholder icons
convert -size 16x16 xc:#0078d4 -fill white -gravity center -pointsize 8 -annotate 0 "R" icon-16.png
convert -size 32x32 xc:#0078d4 -fill white -gravity center -pointsize 16 -annotate 0 "RST" icon-32.png
convert -size 80x80 xc:#0078d4 -fill white -gravity center -pointsize 32 -annotate 0 "RST" icon-80.png
```

Or use an online tool like:
- https://www.favicon-generator.org/
- https://realfavicongenerator.net/
