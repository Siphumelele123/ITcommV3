# 🎨 Blazor UI - Design Reference

Quick visual reference for the styling and layout structure.

## Color Palette Reference

```
PRIMARY ACCENT (Key Actions)
  ├─ Default:  #2B4C6F (Muted Slate Navy)
  └─ Hover:    #1E3A54 (Darker Shade)
	 Used for: "Get ABSA SFTP Files", "Process All Unprocessed Files", etc.

SECONDARY ACTIONS (Utility Buttons)
  ├─ Background: #E9EBEF (Light Gray)
  └─ Text:       #3A3F47 (Dark Gray)
	 Used for: "Get Redo Files", "Clear List", "Check Files", etc.

TEXT & LABELS
  ├─ Primary:    #2A2E35 (Near-Black - Main Content)
  └─ Secondary:  #6B7280 (Medium Gray - Labels & Help Text)

SURFACES & BACKGROUNDS
  ├─ Main BG:    #F5F6F8 (Very Light Cool Gray)
  ├─ Surfaces:   #FFFFFF (White - Cards & Inputs)
  └─ Borders:    #D8DCE3 (Light Gray-Blue)
```

## Layout Structure

### Desktop (>768px)
```
┌─────────────────────────────────────────────────────────┐
│                     Activity Log                        │
├─────────────┬──────────────────────┬───────────────────┤
│   Sidebar   │                      │                   │
│    Left     │   Main Content       │   Right Sidebar   │
│             │                      │   (Advanced Opts) │
│  • Account  │ • Process New Files  │  • File Type      │
│  • Folders  │ • Redo Files         │  • Legacy Files   │
│  • Paths    │ • File Management    │  • Data Tools     │
│             │                      │                   │
│ [Primary]   │   [Primary] [Sec]    │  [Secondary]      │
│ [Primary]   │   [Secondary]        │  [Secondary]      │
└─────────────┴──────────────────────┴───────────────────┘
```

### Tablet (768px - 640px)
```
┌──────────────────────────────────────────────┐
│        Activity Log                          │
├──────────────────────────────────────────────┤
│ Sidebar Left       │  Main Content           │
│                    │                         │
│ • Account          │ • Process New Files     │
│ • Folders          │ • Redo Files            │
│ • Paths            │ • File Management       │
│                    │                         │
│ [Primary] [Primary]│  [Primary] [Secondary]  │
├────────────────────┴─────────────────────────┤
│ Right Sidebar (Advanced Options)            │
│ • File Type                                  │
│ • Legacy Files                               │
│ • Data Tools                                 │
│ [Secondary] [Secondary] [Secondary]         │
└──────────────────────────────────────────────┘
```

### Mobile (<640px)
```
┌──────────────────┐
│  Activity Log    │
├──────────────────┤
│ Sidebar - Left   │
│ • Account        │
│ • Folders        │
│ • Paths          │
│ [Primary]        │
│ [Primary]        │
├──────────────────┤
│ Main Content     │
│ • Process Files  │
│ • Redo Files     │
│ • File Mgmt      │
│ [Primary]        │
│ [Secondary]      │
├──────────────────┤
│ Right Sidebar    │
│ • File Type      │
│ • Legacy Files   │
│ • Data Tools     │
│ [Secondary]      │
│ [Secondary]      │
└──────────────────┘
```

## Button Styling

### Primary Buttons
```
State:    Default           Hover              Active
Color:    #2B4C6F          #1E3A54            Scale 0.98
Text:     White            White              White
Shadow:   None             None               0.98x scale
Border:   None             None               None

Used in:
  • Get Mail Box
  • Get ABSA SFTP Files
  • Get Unprocessed Files
  • Process All Unprocessed Files
```

### Secondary Buttons
```
State:    Default           Hover              Active
BG:       #E9EBEF          #DFE2EA            Scale 0.98
Text:     #3A3F47          #3A3F47            #3A3F47
Border:   #D8DCE3          #D8DCE3            #D8DCE3
Shadow:   None             None               None

Used in:
  • Get Redo Files
  • Mark as Redo
  • Check Files
  • Add to List
  • Clear List
  • All utility actions
```

## Spacing System

```
--spacing-xs:  4px   (Tight spacing, internal margins)
--spacing-sm:  8px   (Small spacing, labels to inputs)
--spacing-md:  16px  (Standard spacing, between elements)
--spacing-lg:  24px  (Large spacing, major sections)
--spacing-xl:  32px  (Extra large, page margins)

Border Radius: 6px (All inputs, buttons, cards)
```

## Responsive Breakpoints

```
Desktop:      > 768px  → 3-column grid layout
Tablet:       768px    → 2-column layout
Mobile:       640px    → 1-column layout
Small Mobile: 480px    → Optimized fonts & spacing
```

## Component Reference

### Card
```html
<div class="card">
	<h3 class="card-title">Section Title</h3>
	<div class="section-group">
		<!-- Content -->
	</div>
</div>
```
- Background: #FFFFFF
- Border: 1px solid #D8DCE3
- Border-radius: 6px
- Padding: 16px
- Box-shadow: 0 1px 3px rgba(0, 0, 0, 0.05)

### Form Group
```html
<div class="form-group">
	<label>Label Text</label>
	<input class="form-control" />
</div>
```
- Label color: #6B7280
- Label size: 13px
- Input padding: 8px 16px
- Input border: 1px solid #D8DCE3
- Focus: Border #2B4C6F + shadow

### List Box / Select
```html
<select class="list-box" size="5"></select>
```
- Same border & focus as form-control
- Multi-select friendly
- Scrollable with styled scrollbar

### Section Group
```html
<div class="section-group">
	<h4 class="section-subtitle">Group Name</h4>
	<!-- Content -->
</div>
```
- Subtitle: Uppercase, 13px, gray
- Margin-bottom: 16px

## Text Styles

```
Headings:
  h1, h2: Never used (use h3+ for cards)
  h3: 16px, weight 600, color #2A2E35
  h4: 13px, weight 600, color #6B7280, uppercase

Body Text:
  Regular:   14px, #2A2E35, line-height 1.5
  Secondary: 14px, #6B7280, line-height 1.5
  Small:     12px, #6B7280

Labels:
  13px, weight 500, #6B7280, uppercase, letter-spacing 0.3px
```

## Transitions

```
All interactive elements: 0.2s ease
  • Button state changes
  • Input focus
  • Hover effects
  • Border color changes

No animation: Kept minimal for professional feel
```

## Hover & Focus States

```
Buttons:
  └─ Background color change + scale 0.98 on active

Form Controls:
  └─ Border color to #2B4C6F
  └─ Box-shadow: 0 0 0 2px rgba(43, 76, 111, 0.1)

Links: (Not primary in this UI)
  └─ Color change + underline
```

---

## Quick Copy-Paste Utilities

Add to any element for quick styling:

```html
<!-- Margin -->
<div class="mt-1">...</div> <!-- 4px margin-top -->
<div class="mt-2">...</div> <!-- 8px -->
<div class="mt-3">...</div> <!-- 16px -->
<div class="mt-4">...</div> <!-- 24px -->

<!-- Same for mb-1, mb-2, mb-3, mb-4 (margin-bottom) -->

<!-- Width -->
<button class="w-100">Full Width</button>

<!-- Text Color -->
<span class="text-muted">Gray text</span>
<span class="text-center">Centered text</span>

<!-- Gap (in flex containers) -->
<div class="gap-1">...</div> <!-- 4px gap -->
<div class="gap-2">...</div> <!-- 8px gap -->
<div class="gap-3">...</div> <!-- 16px gap -->
<div class="gap-4">...</div> <!-- 24px gap -->
```

---

## Design Philosophy

✨ **Muted & Professional**
  - No bright primary colors
  - Soft shadows, not harsh
  - Refined typography
  - Subtle transitions

📱 **Mobile-First Responsive**
  - Starts mobile, expands to desktop
  - Touch-friendly spacing
  - Readable on all sizes
  - No horizontal scrolling

🎯 **Hierarchy & Clarity**
  - Primary actions stand out
  - Secondary actions quieter
  - Clear sections & grouping
  - Proper label-input association

♿ **Accessible & Usable**
  - Semantic HTML
  - Proper contrast ratios
  - Keyboard navigable
  - Focus states visible

---

This design system ensures consistency across the entire Blazor UI!
