<template>
  <div class="editor-shell">
    <AppHeader>
      <div class="file">
        <input v-model="baseName" class="filename" aria-label="File name" spellcheck="false" :style="{ width: `${Math.max(baseName.length, 4) + 2}ch` }" />
        <span class="ext">.docx</span>
        <span v-if="dirty" class="dirty" title="Unsaved changes">●</span>
      </div>
      <button class="primary" type="button" :disabled="saving" title="Download (Ctrl+S)" @click="requestDownload">
        {{ saving ? 'Saving…' : 'Download' }}
      </button>
      <button class="ghost" type="button" @click="close">Close</button>
    </AppHeader>

    <div class="toolbar" :class="{ 'has-changes': changes.total || changes.wordOnly }" role="toolbar" aria-label="Formatting" @mousedown="keepFocus">
      <button type="button" title="Undo (Ctrl+Z)" @click="undo">↶</button>
      <button type="button" title="Redo (Ctrl+Y)" @click="redo">↷</button>
      <span class="sep" />
      <select :value="state.block" aria-label="Paragraph style" title="Paragraph style (Ctrl+Shift+N, Ctrl+Alt+1–3)" @change="applyBlockStyle($event.target.value)">
        <option value="normal">Normal</option>
        <option value="title">Title</option>
        <option value="h1">Heading 1</option>
        <option value="h2">Heading 2</option>
        <option value="h3">Heading 3</option>
        <option v-if="!['normal', 'title', 'h1', 'h2', 'h3'].includes(state.block)" :value="state.block" disabled>
          {{ state.block === 'other' ? 'Other style' : state.block.replace('h', 'Heading ') }}
        </option>
      </select>
      <span class="sep" />
      <button type="button" title="Bold (Ctrl+B)" :class="{ on: state.bold }" @click="exec('bold')"><b>B</b></button>
      <button type="button" title="Italic (Ctrl+I)" :class="{ on: state.italic }" @click="exec('italic')"><i>I</i></button>
      <button type="button" title="Underline (Ctrl+U)" :class="{ on: state.underline }" @click="exec('underline')"><u>U</u></button>
      <button type="button" title="Strikethrough" :class="{ on: state.strike }" @click="exec('strikeThrough')"><s>S</s></button>
      <button type="button" title="Superscript (Ctrl+Shift+=)" :class="{ on: state.sup }" :aria-pressed="state.sup" @click="toggleVertAlign('superscript')"><span class="script">x<sup>2</sup></span></button>
      <button type="button" title="Subscript (Ctrl+=)" :class="{ on: state.sub }" :aria-pressed="state.sub" @click="toggleVertAlign('subscript')"><span class="script">x<sub>2</sub></span></button>
      <span class="split">
        <button type="button" :title="`Highlight: ${colorLabel(highlightColor)} (Ctrl+Alt+H)`" :class="{ on: !!state.highlight }" @click="applyHighlight(highlightColor, true)">
          <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path d="M4 9.5l5.5-6 3 3-6 5.5h-2.5zM4 9.5l-1.5 2.5h3" stroke="currentColor" stroke-width="1.3" stroke-linejoin="round" fill="none" /><rect x="1.5" y="13" width="13" height="2.2" :fill="HIGHLIGHT[highlightColor]" stroke="currentColor" stroke-width=".4" /></svg>
        </button>
        <button ref="paletteButton" type="button" class="arrow" title="Highlight colour" aria-haspopup="dialog" :aria-expanded="palette.open" @click="togglePalette">▾</button>
      </span>
      <span class="sep" />
      <button v-for="a in aligns" :key="a.value" type="button" :title="a.title" :class="{ on: state.align === a.value }" @click="alignBlocks(a.value)">
        <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path :d="a.icon" stroke="currentColor" stroke-width="1.6" fill="none" /></svg>
      </button>
      <span class="sep" />
      <button type="button" title="Bullets (Ctrl+Shift+L)" :class="{ on: state.list === 'bullet' }" :aria-pressed="state.list === 'bullet'" @click="toggleList('bullet')">
        <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path d="M6.5 4h8M6.5 8h8M6.5 12h8M2.5 4h.01M2.5 8h.01M2.5 12h.01" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" fill="none" /></svg>
      </button>
      <button type="button" title="Numbering" :class="{ on: state.list === 'number' }" :aria-pressed="state.list === 'number'" @click="toggleList('number')">
        <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path d="M6.5 4h8M6.5 8h8M6.5 12h8M2 2.5h1.2v3M1.6 5.5h2M1.6 9.6q0-1.6 1.6-1 .9.6-1.6 3h2" stroke="currentColor" stroke-width="1.1" fill="none" /></svg>
      </button>
      <button type="button" title="Decrease list level (Shift+Tab, Ctrl+Shift+M)" :disabled="!state.list" @click="changeLevel(-1)">
        <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path d="M7 4h7M7 8h7M7 12h7M4.5 5.5L2 8l2.5 2.5" stroke="currentColor" stroke-width="1.6" fill="none" /></svg>
      </button>
      <button type="button" title="Increase list level (Tab, Ctrl+M)" :disabled="!state.list" @click="changeLevel(1)">
        <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path d="M7 4h7M7 8h7M7 12h7M2 5.5L4.5 8 2 10.5" stroke="currentColor" stroke-width="1.6" fill="none" /></svg>
      </button>
      <span class="sep" />
      <button type="button" title="Insert picture" @click="pickPicture">
        <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><rect x="1.8" y="2.8" width="12.4" height="10.4" rx="1" stroke="currentColor" stroke-width="1.4" fill="none" /><circle cx="5.6" cy="6.3" r="1.3" fill="currentColor" /><path d="M2.5 12.3l3.8-3.8 2.8 2.8 2-2 2.6 2.6" stroke="currentColor" stroke-width="1.4" stroke-linejoin="round" fill="none" /></svg>
      </button>
      <input ref="pictureInput" type="file" accept="image/png,image/jpeg,image/gif,image/bmp" multiple hidden @change="onPicturesPicked" />
      <button ref="tableButton" type="button" title="Insert table" aria-haspopup="dialog" :aria-expanded="tablePicker.open" @click="toggleTablePicker">
        <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><rect x="1.8" y="2.3" width="12.4" height="11.4" rx="1" stroke="currentColor" stroke-width="1.4" fill="none" /><path d="M1.8 6h12.4M1.8 9.8h12.4M6 2.3v11.4M10 2.3v11.4" stroke="currentColor" stroke-width="1.1" fill="none" /></svg>
      </button>
      <button ref="layoutButton" type="button" title="Table layout: insert or delete rows and columns" aria-haspopup="menu" :aria-expanded="tableMenu.open" :disabled="!state.table" @click="toggleTableMenu">
        <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path d="M9 2.3H2.8a1 1 0 00-1 1v9.4a1 1 0 001 1h9.4a1 1 0 001-1V8" stroke="currentColor" stroke-width="1.4" fill="none" /><path d="M1.8 8h7M6 2.3v11.4" stroke="currentColor" stroke-width="1.1" fill="none" /><path d="M12.5 1.5v5M10 4h5" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" /></svg>
      </button>
      <span class="sep" />
      <button type="button" title="Find and replace (Ctrl+F, Ctrl+H)" :class="{ on: findOpen }" :aria-pressed="findOpen" @click="findOpen ? closeFind() : openFind(false)">
        <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><circle cx="6.5" cy="6.5" r="4.2" stroke="currentColor" stroke-width="1.6" fill="none" /><path d="M9.6 9.6L14 14" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" /></svg>
      </button>
      <button ref="trackButton" type="button" class="track-toggle" title="Track Changes (Ctrl+Shift+E)" :class="{ on: tracking }" :aria-pressed="tracking" @click="toggleTracking">
        <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path d="M2 13.5h5M2 10.5h3M10.5 2.5l3 3-6 6H4.5v-3z" stroke="currentColor" stroke-width="1.5" stroke-linejoin="round" fill="none" /></svg>
        <span class="label">Track changes</span>
      </button>
      <span v-if="changes.total || changes.wordOnly" class="changes-group" role="group" aria-label="Tracked changes">
        <span class="sep" />
        <button type="button" title="Previous change" :disabled="!changes.total" @click="gotoChange(-1)">
          <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path d="M10 3L5 8l5 5" stroke="currentColor" stroke-width="1.6" fill="none" /></svg>
        </button>
        <button type="button" title="Next change" :disabled="!changes.total" @click="gotoChange(1)">
          <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path d="M6 3l5 5-5 5" stroke="currentColor" stroke-width="1.6" fill="none" /></svg>
        </button>
        <span class="split">
          <button type="button" class="labelled" title="Accept the change (or the changes selected) and go to the next" :disabled="!changes.total" @click="resolveCurrent(true)">
            <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path d="M3 8.5l3.2 3.2L13 4.8" stroke="#2e7d32" stroke-width="1.8" fill="none" /></svg><span class="label">Accept</span>
          </button>
          <button ref="acceptMenuButton" type="button" class="arrow" title="More ways to accept" aria-haspopup="menu" :aria-expanded="changeMenu.open === 'accept'" :disabled="!changes.total" @click="openChangeMenu('accept')">▾</button>
        </span>
        <span class="split">
          <button type="button" class="labelled" title="Reject the change (or the changes selected) and go to the next" :disabled="!changes.total" @click="resolveCurrent(false)">
            <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path d="M4 4l8 8M12 4l-8 8" stroke="#b3261e" stroke-width="1.8" fill="none" /></svg><span class="label">Reject</span>
          </button>
          <button ref="rejectMenuButton" type="button" class="arrow" title="More ways to reject" aria-haspopup="menu" :aria-expanded="changeMenu.open === 'reject'" :disabled="!changes.total" @click="openChangeMenu('reject')">▾</button>
        </span>
        <span class="badge" :title="changesTitle">{{ changes.total }}</span>
      </span>
      <button type="button" class="marks-toggle" title="Show formatting marks (Ctrl+Shift+8)" :class="{ on: showMarks }" :aria-pressed="showMarks" @click="toggleMarks">¶</button>
      <button type="button" class="review-toggle" title="Review your changes" @click="openReview">
        <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path d="M2 4l1.2 1.2L5.5 3M2 9l1.2 1.2L5.5 8M7.5 4.2h6.5M7.5 9.2h6.5M7.5 13h6.5" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" fill="none" /></svg>
        <span class="label">Review</span>
      </button>
      <button type="button" class="new-comment" title="New comment (Ctrl+Alt+M)" @click="startComment">
        <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path d="M2.5 3h11v7.5H7L4 13v-2.5H2.5z" stroke="currentColor" stroke-width="1.4" stroke-linejoin="round" fill="none" /><path d="M8 4.8v4M6 6.8h4" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" /></svg>
      </button>
      <button v-if="threads.length" type="button" class="comments-toggle" :title="`Comments (${threads.length})`" :class="{ on: commentsOpen }" :aria-pressed="commentsOpen" @click="commentsOpen = !commentsOpen">
        <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path d="M2.5 3h11v7.5H7L4 13v-2.5H2.5z" stroke="currentColor" stroke-width="1.5" stroke-linejoin="round" fill="none" /></svg>
        <span class="badge">{{ threads.length }}</span>
      </button>
    </div>

    <FindBar v-if="findOpen" ref="findBar" :root="page" :roots="searchRoots" :edit="applyEdit" :tracked="trackedBy" @close="onFindClosed" @reveal="revealRange" />

    <div class="workspace">
      <div class="canvas" @mousedown.self="focusEnd">
        <div class="docx-sheet" :style="sheetStyle" @mousedown.self="focusEnd">
          <div ref="marksEl" class="marks-layer" aria-hidden="true" />
          <div v-if="picture.el" class="picture-frame" :class="{ floating: picture.wrap !== 'inline', cropping: picture.cropping }" :style="{ left: `${picture.x}px`, top: `${picture.y}px`, width: `${picture.w}px`, height: `${picture.h}px` }">
            <span v-for="h in HANDLES" :key="h" class="handle" :class="h" :title="picture.cropping ? 'Drag to crop' : h.length === 2 ? 'Drag to resize' : 'Drag to stretch'" @pointerdown.prevent.stop="picture.cropping ? startCrop($event, h) : startResize($event, h)" />
            <span v-if="!picture.cropping" class="handle rotate" title="Drag to rotate (Shift: in steps of 15°)" @pointerdown.prevent.stop="startRotate" />
          </div>
          <button v-if="picture.el" ref="layoutOptionsButton" type="button" class="layout-options" title="Layout Options: wrapping, position, size and alt text" aria-haspopup="dialog" :aria-expanded="picturePanel.open" :style="{ left: `${picture.x + picture.w + 6}px`, top: `${picture.y}px` }" @mousedown.prevent @click="togglePicturePanel">
            <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><rect x="2" y="3" width="6" height="5" stroke="currentColor" stroke-width="1.3" fill="none" /><path d="M10 4h4M10 7h4M2 11h12M2 14h12" stroke="currentColor" stroke-width="1.3" /></svg>
          </button>
          <div v-if="doc.headers.length" class="docx-hf header" :style="headerStyle" :title="hfTitle('header')" @click="onHeaderFooterClick">
            <div v-if="doc.headers.length > 1" class="hf-tabs" role="tablist" aria-label="Headers">
              <button v-for="(v, i) in doc.headers" :key="v.rid" type="button" role="tab" :aria-selected="hf.header === i" @click="hf.header = i">{{ v.label }}</button>
            </div>
            <div ref="headerEl" class="hf-variants" />
          </div>
          <div
            ref="page"
            class="docx-page"
            contenteditable="true"
            spellcheck="true"
            :style="pageStyle"
            @beforeinput="onBeforeInput"
            @input="onInput"
            @keydown="onKeydown"
            @paste="onPaste"
            @drop="onDrop"
            @mousedown="onPageMousedown"
            @compositionend="onCompositionEnd"
            @click="onClick"
          />
          <div v-if="doc.footers.length" class="docx-hf footer" :style="footerStyle" :title="hfTitle('footer')" @click="onHeaderFooterClick">
            <div ref="footerEl" class="hf-variants" />
            <div v-if="doc.footers.length > 1" class="hf-tabs" role="tablist" aria-label="Footers">
              <button v-for="(v, i) in doc.footers" :key="v.rid" type="button" role="tab" :aria-selected="hf.footer === i" @click="hf.footer = i">{{ v.label }}</button>
            </div>
          </div>
        </div>
      </div>
      <CommentsPanel
        v-if="threads.length || composer.open"
        v-show="commentsOpen || composer.open"
        :threads="threads"
        :active="activeComment"
        :composing="composer.open"
        :ask-name="!named"
        @select="selectComment"
        @close="commentsOpen = false; cancelComment()"
        @post="postComment"
        @cancel="cancelComment"
        @delete="deleteComment"
        @reply="replyComment"
        @edit="editComment"
        @resolve="resolveComment"
      />
    </div>

    <div v-if="chooser.open" ref="chooserEl" class="chooser" :role="chooser.kind === 'date' ? 'dialog' : 'listbox'" aria-label="Choose a value" :style="{ left: `${chooser.x}px`, top: `${chooser.y}px` }" @keydown.esc="closeChooser(true)" @keydown="onChooserKeydown">
      <input v-if="chooser.kind === 'date'" ref="dateInput" type="date" :value="chooser.date" aria-label="Date" @keydown="onDateKeydown" @change="onDateChange" />
      <template v-else>
        <button v-for="c in chooser.choices" :key="c.value" type="button" role="option" :aria-selected="c.text === chooser.current" @click="choose(c)">{{ c.text }}</button>
        <p v-if="!chooser.choices.length" class="empty">No choices</p>
      </template>
    </div>

    <div v-if="changeMenu.open" ref="changeMenuEl" class="change-menu" role="menu" :aria-label="changeMenu.open === 'accept' ? 'Accept' : 'Reject'" :style="{ left: `${changeMenu.x}px`, top: `${changeMenu.y}px` }" @mousedown.prevent @keydown.esc="closeChangeMenu">
      <template v-if="changeMenu.open === 'accept'">
        <button type="button" role="menuitem" @click="menuResolve(true, 'one')">Accept this change</button>
        <button type="button" role="menuitem" @click="menuResolve(true, 'stay')">Accept, and stay here</button>
        <button type="button" role="menuitem" @click="menuResolve(true, 'all')">Accept all changes</button>
      </template>
      <template v-else>
        <button type="button" role="menuitem" @click="menuResolve(false, 'one')">Reject this change</button>
        <button type="button" role="menuitem" @click="menuResolve(false, 'stay')">Reject, and stay here</button>
        <button type="button" role="menuitem" @click="menuResolve(false, 'all')">Reject all changes</button>
      </template>
    </div>

    <form v-if="picturePanel.open" ref="picturePanelEl" class="picture-panel" role="dialog" aria-label="Layout Options" :style="{ left: `${picturePanel.x}px`, top: `${picturePanel.y}px` }" @submit.prevent @keydown.esc="closePicturePanel">
      <fieldset class="wraps">
        <legend>Text wrapping</legend>
        <label v-for="[value, label] in WRAPS" :key="value"><input type="radio" name="wrap" :value="value" :checked="picturePanel.wrap === value" @change="applyWrap(value)" /> {{ label }}</label>
      </fieldset>
      <fieldset :disabled="picturePanel.wrap === 'inline'">
        <legend>Position</legend>
        <label>Horizontal
          <select v-model="picturePanel.hAlign" @change="applyPosition('h')">
            <option value="left">Left of the margins</option>
            <option value="center">Centred between the margins</option>
            <option value="right">Right of the margins</option>
            <option value="offset">At a distance from the left margin</option>
            <option v-if="picturePanel.hAlign === 'word'" value="word">As set in Word</option>
          </select>
        </label>
        <label v-if="picturePanel.hAlign === 'offset'">Distance from the margin (cm) <input v-model.number="picturePanel.hOffset" type="number" step="0.1" @change="applyPosition('h')" /></label>
        <label>Vertical
          <select v-model="picturePanel.vAlign" @change="applyPosition('v')">
            <option value="offset">Below its paragraph (moves with text)</option>
            <option value="top">Top of the margins</option>
            <option value="center">Middle of the margins</option>
            <option value="bottom">Bottom of the margins</option>
            <option v-if="picturePanel.vAlign === 'word'" value="word">As set in Word</option>
          </select>
        </label>
        <label v-if="picturePanel.vAlign === 'offset'">Distance below the paragraph (cm) <input v-model.number="picturePanel.vOffset" type="number" step="0.1" @change="applyPosition('v')" /></label>
      </fieldset>
      <fieldset>
        <legend>Size</legend>
        <div class="size-row">
          <label>Width (cm) <input v-model.number="picturePanel.width" type="number" min="0.1" step="0.1" @change="applySize('width')" /></label>
          <label>Height (cm) <input v-model.number="picturePanel.height" type="number" min="0.1" step="0.1" @change="applySize('height')" /></label>
        </div>
        <label class="inline"><input v-model="picturePanel.lock" type="checkbox" /> Keep proportions</label>
        <button type="button" @click="resetPictureSize">Reset size</button>
      </fieldset>
      <details>
        <summary>Crop</summary>
        <div class="size-row">
          <label>Left (%) <input v-model.number="picturePanel.cropL" type="number" step="1" @change="applyCrop" /></label>
          <label>Right (%) <input v-model.number="picturePanel.cropR" type="number" step="1" @change="applyCrop" /></label>
          <label>Top (%) <input v-model.number="picturePanel.cropT" type="number" step="1" @change="applyCrop" /></label>
          <label>Bottom (%) <input v-model.number="picturePanel.cropB" type="number" step="1" @change="applyCrop" /></label>
        </div>
        <button type="button" :aria-pressed="picture.cropping" @click="toggleCropping">{{ picture.cropping ? 'Done cropping' : 'Crop with the handles' }}</button>
      </details>
      <details>
        <summary>Rotate and flip</summary>
        <div class="buttons">
          <button type="button" @click="rotateBy(90)">Rotate right 90°</button>
          <button type="button" @click="rotateBy(-90)">Rotate left 90°</button>
          <button type="button" :aria-pressed="picturePanel.flipH" @click="flip('flipH')">Flip horizontal</button>
          <button type="button" :aria-pressed="picturePanel.flipV" @click="flip('flipV')">Flip vertical</button>
        </div>
        <label>Rotation (°) <input v-model.number="picturePanel.rot" type="number" step="1" @change="applyRotation" /></label>
      </details>
      <details>
        <summary>Effects</summary>
        <div class="size-row">
          <label>Border
            <select v-model="picturePanel.border" @change="applyEffects">
              <option value="none">No border</option>
              <option value="thin">Thin</option>
              <option value="medium">Medium</option>
              <option value="thick">Thick</option>
            </select>
          </label>
          <label>Border colour <input v-model="picturePanel.borderColor" type="color" @change="applyEffects" /></label>
        </div>
        <label class="inline"><input v-model="picturePanel.shadow" type="checkbox" @change="applyEffects" /> Shadow</label>
        <label class="inline"><input v-model="picturePanel.glow" type="checkbox" @change="applyEffects" /> Glow</label>
        <label class="inline"><input v-model="picturePanel.softEdge" type="checkbox" @change="applyEffects" /> Soft edges</label>
        <label class="inline"><input v-model="picturePanel.reflection" type="checkbox" @change="applyEffects" /> Reflection</label>
        <label>Colour
          <select v-model="picturePanel.colour" @change="applyEffects">
            <option value="normal">As it is</option>
            <option value="greyscale">Greyscale</option>
          </select>
        </label>
        <div class="size-row">
          <label>Brightness (%) <input v-model.number="picturePanel.brightness" type="number" min="-100" max="100" step="5" @change="applyEffects" /></label>
          <label>Contrast (%) <input v-model.number="picturePanel.contrast" type="number" min="-100" max="100" step="5" @change="applyEffects" /></label>
        </div>
      </details>
      <label class="alt">Alt text (describe the picture for people who can't see it)
        <textarea v-model="picturePanel.alt" rows="2" @change="applyAlt" />
      </label>
      <div class="actions"><button type="button" @click="closePicturePanel">Close</button></div>
    </form>

    <div v-if="tableMenu.open" ref="tableMenuEl" class="change-menu" role="menu" aria-label="Table layout" :style="{ left: `${tableMenu.x}px`, top: `${tableMenu.y}px` }" @mousedown.prevent @keydown.esc="closeTableMenu" @keydown="onMenuArrows">
      <button type="button" role="menuitem" @click="tableAction('rowsAbove')">{{ tableMenu.rows > 1 ? `Insert ${tableMenu.rows} rows above` : 'Insert row above' }}</button>
      <button type="button" role="menuitem" @click="tableAction('rowsBelow')">{{ tableMenu.rows > 1 ? `Insert ${tableMenu.rows} rows below` : 'Insert row below' }}</button>
      <button type="button" role="menuitem" @click="tableAction('columnsLeft')">{{ tableMenu.newCols > 1 ? `Insert ${tableMenu.newCols} columns to the left` : 'Insert column to the left' }}</button>
      <button type="button" role="menuitem" @click="tableAction('columnsRight')">{{ tableMenu.newCols > 1 ? `Insert ${tableMenu.newCols} columns to the right` : 'Insert column to the right' }}</button>
      <hr />
      <button type="button" role="menuitem" :disabled="!tableMenu.multi" :title="tableMenu.multi ? '' : 'Select the cells to merge first'" @click="tableAction('merge')">Merge cells</button>
      <button type="button" role="menuitem" @click="openSplitDialog">Split cells…</button>
      <hr />
      <button type="button" role="menuitem" @click="tableAction('deleteRows')">{{ tableMenu.rows > 1 ? `Delete ${tableMenu.rows} rows` : 'Delete row' }}</button>
      <button type="button" role="menuitem" @click="tableAction('deleteColumns')">{{ tableMenu.cols > 1 ? `Delete ${tableMenu.cols} columns` : 'Delete column' }}</button>
      <button type="button" role="menuitem" @click="tableAction('deleteTable')">Delete table</button>
    </div>

    <form v-if="splitDialog.open" ref="splitDialogEl" class="chooser author-dialog split-dialog" role="dialog" aria-label="Split cells" :style="{ left: `${splitDialog.x}px`, top: `${splitDialog.y}px` }" @submit.prevent="splitCells" @keydown.esc="closeSplitDialog">
      <label>Number of columns <input v-model.number="splitDialog.cols" type="number" min="1" max="20" required /></label>
      <label>Number of rows <input v-model.number="splitDialog.rows" type="number" min="1" max="50" required /></label>
      <div class="actions">
        <button type="submit" class="confirm">Split</button>
        <button type="button" @click="closeSplitDialog">Cancel</button>
      </div>
    </form>

    <div v-if="tablePicker.open" ref="tablePickerEl" class="table-picker" role="dialog" aria-label="Insert table" :style="{ left: `${tablePicker.x}px`, top: `${tablePicker.y}px` }" @keydown.esc="closeTablePicker">
      <p class="size" aria-hidden="true">{{ tablePicker.cols }} × {{ tablePicker.rows }} table</p>
      <div class="grid" @mouseleave="tablePicker.cols = tablePicker.rows = 1">
        <template v-for="r in PICKER_ROWS" :key="r">
          <button
            v-for="c in PICKER_COLS"
            :key="c"
            type="button"
            class="cell"
            :class="{ on: r <= tablePicker.rows && c <= tablePicker.cols }"
            :tabindex="r === tablePicker.rows && c === tablePicker.cols ? 0 : -1"
            :aria-label="`${c} by ${r} table`"
            @mousedown.prevent
            @mouseenter="Object.assign(tablePicker, { rows: r, cols: c })"
            @keydown="onPickerKeydown"
            @click="insertTable(r, c)"
          />
        </template>
      </div>
      <form class="custom" @submit.prevent="insertTable(tablePicker.customRows, tablePicker.customCols)">
        <label>Columns <input v-model.number="tablePicker.customCols" type="number" min="1" :max="MAX_COLS" required /></label>
        <label>Rows <input v-model.number="tablePicker.customRows" type="number" min="1" :max="MAX_ROWS" required /></label>
        <button type="submit">Insert</button>
      </form>
    </div>

    <div v-if="palette.open" ref="paletteEl" class="palette" role="dialog" aria-label="Highlight colour" :style="{ left: `${palette.x}px`, top: `${palette.y}px` }" @mousedown.prevent @keydown.esc="closePalette">
      <div class="swatches">
        <button v-for="[name, label] in HIGHLIGHT_COLORS" :key="name" type="button" class="swatch" :class="{ current: name === highlightColor }" :style="{ background: HIGHLIGHT[name] }" :title="label" :aria-label="label" @click="pickHighlight(name)" />
      </div>
      <button type="button" class="no-color" @click="pickHighlight(null)">No colour</button>
    </div>

    <ReviewDialog v-if="reviewData" :review="{ ...reviewData, resolved }" :skip="skipReview" @close="closeReview" @download="downloadReviewed" @skip="setSkipReview" />

    <form v-if="authorDialog.open" ref="authorEl" class="chooser author-dialog" role="dialog" aria-label="Your name on tracked changes" :style="{ left: `${authorDialog.x}px`, top: `${authorDialog.y}px` }" @submit.prevent="confirmAuthor" @keydown.esc="closeAuthorDialog">
      <label for="author-name">Your name on tracked changes</label>
      <!-- No browser autocomplete: Firefox's list of earlier entries opens over the buttons below and
           takes the click meant for them. The name is remembered by Intact itself. -->
      <input id="author-name" ref="authorInput" v-model="authorDialog.name" type="text" maxlength="100" autocomplete="off" required />
      <div class="actions">
        <button type="submit" class="confirm">{{ tracking ? 'Save' : 'Start tracking' }}</button>
        <button type="button" @click="closeAuthorDialog">Cancel</button>
      </div>
    </form>

    <footer class="statusbar">
      <span>{{ words }} words</span>
      <button v-if="tracking" type="button" class="tracking-status" title="Change your name on tracked changes" @click="openAuthorDialog">Tracking changes as {{ author }}</button>
      <span role="status" aria-live="polite">{{ message }}</span>
    </footer>
  </div>
</template>

<script setup>
import { ref, reactive, computed, watch, nextTick, onMounted, onBeforeUnmount } from 'vue';
import { useRouter, onBeforeRouteLeave } from 'vue-router';
import AppHeader from '../components/AppHeader.vue';
import FindBar from '../components/FindBar.vue';
import CommentsPanel from '../components/CommentsPanel.vue';
import ReviewDialog from '../components/ReviewDialog.vue';
import { currentDoc, setDocument } from '../store.js';
import { createHistory } from '../editor/history.js';
import { isEmptyBlock, isAtBlockStart, adjacentCell, selectCell } from '../editor/structure.js';
import { controlAt, controlBox, setControlText, clearPlaceholder, fieldAtSelection } from '../editor/controls.js';
import { parseClipboard, runNodes, insertAtCaret } from '../editor/paste.js';
import {
  removesContent, rangesRemovedBy, trackedChangeAt, refusalMessage, commentMarkers, restoreCommentMarkers,
  onlyHiddenMarkers, beyondHiddenMarkers,
} from '../editor/protect.js';
import { readOnlyAt, readOnlyMessage, isEditingKey, caretBeside } from '../editor/readonly.js';
import { shortcutFor } from '../editor/shortcuts.js';
import { HIGHLIGHT_COLORS, setVertAlign, setHighlight, formatState } from '../editor/inline.js';
import { drawMarks } from '../editor/marks.js';
import { changeItems, changesAt, adjacentChange, itemRange, mergeGroup } from '../editor/changes.js';
import { HIGHLIGHT } from '../docx/styles.js';
import { WRAPS } from '../docx/pictures.js';
import { pictureOutline } from '../editor/outline.js';
import {
  revisionDate, typeTracked, typeUntracked, insertTracked, deleteTracked, unitRange, markSplit, wrapInsertedBreak, wrapComposedText,
  inField,
} from '../editor/track.js';

const router = useRouter();
const doc = currentDoc.value;
const page = ref(null);
const dirty = ref(false);
const saving = ref(false);
const message = ref('');
const words = ref(0);
const baseName = ref(doc.name.replace(/\.docx$/i, ''));
const state = reactive({ bold: false, italic: false, underline: false, strike: false, sup: false, sub: false, highlight: null, block: 'normal', align: 'left', list: null, table: false });
const chooser = reactive({ open: false, kind: '', x: 0, y: 0, choices: [], current: '', date: '' });
const chooserEl = ref(null);
const threads = ref([]);
const commentsOpen = ref(false);
const activeComment = ref(null);
const findOpen = ref(false);

// Track Changes: on when the document says so (as in Word), or when switched on here. The author's
// name is remembered in this browser only.
const AUTHOR_KEY = 'intact.author';
const stored = (() => { try { return localStorage.getItem(AUTHOR_KEY) || ''; } catch { return ''; } })();
const tracking = ref(doc.tracking);
const author = ref(stored || 'Author');
const named = ref(!!stored); // the user has given a name
const authorDialog = reactive({ open: false, name: '', x: 0, y: 0, start: false });
const trackButton = ref(null);
const authorEl = ref(null);
const authorInput = ref(null);
watch(tracking, (on) => { doc.tracking = on; });
const findBar = ref(null);
const dateInput = ref(null);
let chooserTarget = null;
let savedRange = null;
let styleEl = null;
let undoStack = null;

const BLOCK_SELECTOR = 'p,h1,h2,h3,h4,h5,h6,li';
const aligns = [
  { value: 'left', title: 'Align left (Ctrl+L)', icon: 'M2 3h12M2 6.5h8M2 10h12M2 13.5h8' },
  { value: 'center', title: 'Center (Ctrl+E)', icon: 'M2 3h12M4 6.5h8M2 10h12M4 13.5h8' },
  { value: 'right', title: 'Align right (Ctrl+R)', icon: 'M2 3h12M6 6.5h8M2 10h12M6 13.5h8' },
  { value: 'justify', title: 'Justify (Ctrl+J)', icon: 'M2 3h12M2 6.5h12M2 10h12M2 13.5h12' },
];

// The sheet is the page: header (in the top margin), the editable body, footer (bottom margin). As in
// Word, the body starts at the top margin or below the header, whichever is lower.
const headerEl = ref(null);
const footerEl = ref(null);
const sheetStyle = computed(() => ({ width: `${doc.page.width}px`, minHeight: `${doc.page.minHeight}px` }));
const pageStyle = computed(() => {
  const p = doc.page;
  return { padding: `${doc.headers.length ? 0 : p.top}px ${p.right}px ${doc.footers.length ? 0 : p.bottom}px ${p.left}px` };
});
const headerStyle = computed(() => {
  const p = doc.page;
  return { padding: `${p.header}px ${p.right}px 0 ${p.left}px`, minHeight: `${Math.max(0, p.top - p.header)}px` };
});
const footerStyle = computed(() => {
  const p = doc.page;
  return { padding: `0 ${p.right}px ${p.footer}px ${p.left}px`, minHeight: `${Math.max(0, p.bottom - p.footer)}px` };
});

// Which header / footer is shown, when the document has several (first page, even pages, sections).
const hf = reactive({ header: 0, footer: 0 });

function hfTitle(which) {
  const variant = (which === 'header' ? doc.headers : doc.footers)[hf[which]];
  const label = variant?.label || '';
  const name = label === 'All pages' || !label ? which : `${label} ${which}`;
  return `${name[0].toUpperCase()}${name.slice(1)} – shown for reference; edit it in Word`;
}

/** Renders every header (or footer) variant once; only the selected one is visible. */
function renderVariants(el, which) {
  if (!el) return;
  const variants = which === 'header' ? doc.headers : doc.footers;
  const nodes = variants.map((v, i) => {
    const div = document.createElement('div');
    div.className = 'hf-variant';
    div.dataset.which = which;
    div.dataset.variant = String(i);
    div.hidden = i !== hf[which];
    const content = doc.renderHeaderFooter(v);
    if (content) div.append(...content.childNodes);
    return div;
  });
  el.replaceChildren(...nodes);
  el.parentElement.classList.toggle('empty', variants.length === 1 && !nodes[0].childNodes.length);
}

watch(() => ({ ...hf }), () => {
  for (const el of [headerEl.value, footerEl.value]) {
    for (const v of el?.children || []) v.hidden = Number(v.dataset.variant) !== hf[v.dataset.which];
  }
});

function refreshCss() {
  styleEl.textContent = doc.css('.docx-sheet');
}

onMounted(() => {
  styleEl = document.createElement('style');
  document.head.append(styleEl);
  refreshCss();
  page.value.replaceChildren(doc.render());
  renderVariants(headerEl.value, 'header');
  renderVariants(footerEl.value, 'footer');
  threads.value = commentThreads();
  commentsOpen.value = threads.value.length > 0 && window.innerWidth >= 1100;
  paintComments();
  document.execCommand('defaultParagraphSeparator', false, 'p');
  document.execCommand('styleWithCSS', false, false);
  document.addEventListener('selectionchange', onSelectionChange);
  document.addEventListener('mousedown', onDocumentMousedown, true);
  window.addEventListener('keydown', onWindowKeydown, true);
  window.addEventListener('beforeunload', onBeforeUnload);
  countWords();
  refreshChanges();
  page.value.focus();
  placeCaret(page.value.querySelector(BLOCK_SELECTOR) || page.value, true);
  undoStack = createHistory(page.value);
  undoStack.commit();
});

// ------------------------------------------------------------------ formatting marks

// Shown or not is remembered in this browser, as Word remembers it.
const MARKS_KEY = 'intact.marks';
const showMarks = ref((() => { try { return localStorage.getItem(MARKS_KEY) === '1'; } catch { return false; } })());
const marksEl = ref(null);
let marksFrame = 0;
let marksObserver = null;

/** Redraws the formatting marks on the next frame (once, however often it is asked). */
function scheduleMarks() {
  if (!showMarks.value || marksFrame) return;
  marksFrame = requestAnimationFrame(() => {
    marksFrame = 0;
    if (showMarks.value && page.value && marksEl.value) drawMarks(page.value, marksEl.value, page.value.closest('.canvas').getBoundingClientRect());
  });
}

function toggleMarks() {
  showMarks.value = !showMarks.value;
  try { localStorage.setItem(MARKS_KEY, showMarks.value ? '1' : '0'); } catch { /* this session only */ }
  watchMarks();
}

/** Keeps the marks up to date while they are shown: the page changing, scrolling or resizing. */
function watchMarks() {
  const canvas = page.value?.closest('.canvas');
  if (showMarks.value) {
    marksObserver ||= new MutationObserver(scheduleMarks);
    marksObserver.observe(page.value, { childList: true, subtree: true, characterData: true, attributes: true });
    canvas?.addEventListener('scroll', scheduleMarks, { passive: true });
    window.addEventListener('resize', scheduleMarks);
    scheduleMarks();
  } else {
    marksObserver?.disconnect();
    canvas?.removeEventListener('scroll', scheduleMarks);
    window.removeEventListener('resize', scheduleMarks);
    marksEl.value?.replaceChildren();
  }
}

onMounted(() => {
  if (showMarks.value) watchMarks();
  document.fonts?.ready.then(scheduleMarks); // web fonts change where the text is
});

onBeforeUnmount(() => {
  marksObserver?.disconnect();
  cancelAnimationFrame(marksFrame);
  window.removeEventListener('resize', scheduleMarks);

  document.removeEventListener('selectionchange', onSelectionChange);
  document.removeEventListener('mousedown', onDocumentMousedown, true);
  window.removeEventListener('keydown', onWindowKeydown, true);
  window.removeEventListener('beforeunload', onBeforeUnload);
  undoStack?.dispose();
  styleEl?.remove();
  if (canHighlight) {
    CSS.highlights.delete('comment');
    CSS.highlights.delete('comment-active');
    CSS.highlights.delete('field-near');
  }
});

onBeforeRouteLeave(() => {
  if (dirty.value && !window.confirm('You have unsaved changes. Leave without downloading?')) return false;
  return true;
});

function onBeforeUnload(e) {
  if (dirty.value) e.preventDefault();
}

// ------------------------------------------------------------------ selection & toolbar state

function selectionInPage() {
  const sel = window.getSelection();
  if (!sel.rangeCount) return null;
  const range = sel.getRangeAt(0);
  return page.value?.contains(range.commonAncestorContainer) ? range : null;
}

/**
 * The paragraph a selection starts in. Firefox starts a Select All at the page itself (offset 0)
 * rather than inside the first paragraph, so look down into the node the offset points at.
 */
function startBlock(range) {
  const { startContainer: c, startOffset: o } = range;
  const at = c.nodeType === 1 ? c.childNodes[o] || c : c;
  return closestBlock(at) || (at.nodeType === 1 ? at.querySelector(BLOCK_SELECTOR) : null);
}

function closestBlock(node) {
  const el = node.nodeType === 1 ? node : node.parentElement;
  const block = el?.closest(BLOCK_SELECTOR);
  return block && page.value.contains(block) ? block : null;
}

function onSelectionChange() {
  const range = selectionInPage();
  if (!range) return;
  lastReadOnly = null;
  const pic = pictureInSelection(range);
  if (pic !== picture.el) {
    if (pic) showPicture(pic);
    else if (!picturePanel.open) picture.el = null;
  }
  // Clicking into placeholder text selects all of it, so typing replaces it (as in Word).
  const control = range.collapsed ? controlAt(range.startContainer, page.value) : null;
  if (control?.classList.contains('showing-placeholder') && control.querySelector(':scope > .sdt-content')) {
    selectCell(controlBox(control));
    return;
  }
  savedRange = range.cloneRange();
  markNearField(range);
  state.bold = document.queryCommandState('bold');
  state.italic = document.queryCommandState('italic');
  state.underline = document.queryCommandState('underline');
  state.strike = document.queryCommandState('strikeThrough');
  Object.assign(state, (({ superscript, subscript, highlight }) => ({ sup: superscript, sub: subscript, highlight }))(formatState(range)));
  const block = startBlock(range);
  state.block = block ? doc.blockKind(block.dataset.style) : 'normal';
  const align = block ? getComputedStyle(block).textAlign : 'left';
  state.align = { start: 'left', end: 'right' }[align] || align;
  state.list = block ? doc.listKind(doc.listOf(block)) : null;
  state.table = !!tableSelection();
}

/** Back to the page with the selection it had (after a menu, a dialog or the file picker). */
function refocus() {
  const range = savedRange;
  page.value.focus();
  if (range && page.value.contains(range.commonAncestorContainer)) {
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
  }
}

function restoreSelection() {
  if (selectionInPage()) return;
  page.value.focus();
  if (savedRange) {
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(savedRange);
  }
}

function keepFocus(e) {
  // Keep the caret in the document when clicking toolbar buttons (but let <select> open).
  if (e.target.closest('button')) e.preventDefault();
}

function placeCaret(el, atStart = false) {
  const range = document.createRange();
  range.selectNodeContents(el);
  range.collapse(atStart);
  const sel = window.getSelection();
  sel.removeAllRanges();
  sel.addRange(range);
}

function focusEnd() {
  page.value.focus();
  const blocks = page.value.querySelectorAll(':scope > ' + BLOCK_SELECTOR.split(',').join(', :scope > '));
  placeCaret(blocks[blocks.length - 1] || page.value);
}

// ------------------------------------------------------------------ editing commands

const CHAR_FORMATS = new Set(['bold', 'italic', 'underline', 'strikeThrough']);

/** Plain-text fields (as in Word) take no character formatting. */
function inPlainField() {
  const range = selectionInPage();
  const box = range && controlAt(range.startContainer, page.value)?.querySelector(':scope > [data-plain]');
  if (box) message.value = 'Plain text fields can’t be formatted';
  return !!box;
}

function exec(cmd, value = null) {
  restoreSelection();
  if (CHAR_FORMATS.has(cmd) && inPlainField()) return;
  undoStack.flush();
  document.execCommand(cmd, false, value);
  undoStack.commit();
  onSelectionChange();
  markDirty();
  noteUntrackedFormatting();
}

/**
 * Aligns the selected paragraphs. Done here rather than with the browser's justify commands, which
 * Firefox carries out with an align attribute instead of the text-align style that is saved.
 */
function alignBlocks(value) {
  restoreSelection();
  const range = selectionInPage();
  if (!range) return;
  undoStack.flush();
  for (const block of selectedBlocks(range)) {
    block.style.textAlign = value;
    block.removeAttribute('align');
  }
  undoStack.commit();
  onSelectionChange();
  markDirty();
  noteUntrackedFormatting();
}

function afterHistory(changed) {
  if (!changed) return;
  savedRange = null;
  refreshChanges();
  refreshComments(); // undo/redo replaced the DOM the comment ranges pointed into
  onSelectionChange();
  markDirty();
}
const undo = () => { closeChooser(); afterHistory(undoStack.undo()); };
const redo = () => { closeChooser(); afterHistory(undoStack.redo()); };

function selectedBlocks(range) {
  const all = [...page.value.querySelectorAll(BLOCK_SELECTOR)].filter((b) => range.intersectsNode(b) && !b.closest('.sdt-locked'));
  if (all.length) return all;
  const one = closestBlock(range.startContainer);
  return one ? [one] : [];
}

function applyBlockStyle(kind) {
  restoreSelection();
  const range = selectionInPage();
  if (!range) return;
  undoStack.flush();
  const styleId = doc.styleIdFor(kind);
  refreshCss(); // a missing builtin style may just have been added
  const tag = kind === 'normal' ? 'p' : kind === 'title' ? 'h1' : kind;
  const { startContainer, startOffset, endContainer, endOffset } = range;
  const replaced = new Map();
  for (const block of selectedBlocks(range)) {
    let el = block;
    if (block.tagName.toLowerCase() !== tag) {
      el = document.createElement(tag);
      for (const attr of block.attributes) el.setAttribute(attr.name, attr.value);
      while (block.firstChild) el.append(block.firstChild);
      block.replaceWith(el);
      replaced.set(block, el);
    }
    if (styleId) el.dataset.style = styleId;
    else delete el.dataset.style;
  }
  const r = document.createRange();
  r.setStart(replaced.get(startContainer) || startContainer, startOffset);
  r.setEnd(replaced.get(endContainer) || endContainer, endOffset);
  const sel = window.getSelection();
  sel.removeAllRanges();
  sel.addRange(r);
  doc.renumber(page.value);
  undoStack.commit();
  onSelectionChange();
  markDirty();
  noteUntrackedFormatting();
}

// ------------------------------------------------------------------ lists

/** Runs a change to list membership as one undoable step. */
function changeLists(change) {
  restoreSelection();
  const range = selectionInPage();
  if (!range) return;
  undoStack.flush();
  if (change(selectedBlocks(range)) === false) return;
  doc.renumber(page.value);
  undoStack.commit();
  onSelectionChange();
  markDirty();
  noteUntrackedFormatting();
}

/** The list next to `block` (before or after it), if it is of `kind`. */
function adjacentList(block, before, kind) {
  let el = before ? block.previousElementSibling : block.nextElementSibling;
  while (el?.classList.contains('raw-hidden')) el = before ? el.previousElementSibling : el.nextElementSibling;
  const list = el?.matches(BLOCK_SELECTOR) ? doc.listOf(el) : null;
  return list && doc.listKind(list) === kind ? list : null;
}

/**
 * Bullets / Numbering, as in Word: turns the selected paragraphs into a list of that kind (joining
 * a list of the same kind right before or after them), or takes them out of it if they already are.
 */
function toggleList(kind) {
  changeLists((blocks) => {
    if (!blocks.length) return false;
    const lists = blocks.map((b) => doc.listOf(b));
    if (lists.every((l) => doc.listKind(l) === kind)) {
      blocks.forEach((b) => doc.removeList(b));
      return true;
    }
    const join = adjacentList(blocks[0], true, kind) || adjacentList(blocks[blocks.length - 1], false, kind);
    const numId = join?.numId ?? doc.newList(kind);
    blocks.forEach((b, i) => {
      const ilvl = lists[i]?.ilvl ?? join?.ilvl ?? 0;
      doc.setList(b, numId, doc.hasListLevel(numId, ilvl) ? ilvl : 0);
    });
    return true;
  });
}

/** Moves the selected list items one level down (+1) or up (-1). */
function changeLevel(delta) {
  changeLists((blocks) => {
    const items = blocks.map((b) => [b, doc.listOf(b)]).filter(([, l]) => l && doc.hasListLevel(l.numId, l.ilvl + delta));
    if (!items.length) {
      if (delta > 0 && blocks.some((b) => doc.listOf(b))) message.value = 'This list has no further levels';
      return false;
    }
    items.forEach(([b, l]) => doc.setList(b, l.numId, l.ilvl + delta));
    return true;
  });
}

/** Enter on an empty list item moves it up a level, and out of the list from the top level. */
function endListItem(block) {
  const list = doc.listOf(block);
  changeLists(() => {
    if (list.ilvl > 0) doc.setList(block, list.numId, list.ilvl - 1);
    else doc.removeList(block);
  });
}

/** Tab / Shift+Tab: next or previous table cell, else list level, else a tab character. */
function onTab(e) {
  const range = selectionInPage();
  if (!range) return;
  const start = range.startContainer;
  const cell = (start.nodeType === 1 ? start : start.parentElement).closest('td');
  if (cell && page.value.contains(cell)) {
    const target = adjacentCell(cell, e.shiftKey ? -1 : 1);
    if (target) selectCell(target);
    else if (!e.shiftKey) addRow(cell); // Tab in the last cell, as in Word
    return;
  }
  const block = closestBlock(range.startContainer);
  const inList = range.collapsed
    ? block && doc.listOf(block) && isAtBlockStart(range, block)
    : selectedBlocks(range).some((b) => doc.listOf(b));
  if (inList) changeLevel(e.shiftKey ? -1 : 1);
  else if (!e.shiftKey) {
    const range = selectionInPage();
    if (range && !range.collapsed && refuseRemoving([range])) return;
    if (range && !range.collapsed) deleteSelection();
    document.execCommand('insertText', false, '\t');
  }
}

// ------------------------------------------------------------------ events

let wordTimer = 0;
function countWords() {
  const text = page.value?.innerText || '';
  words.value = (text.match(/[\p{L}\p{N}]+(?:['’-][\p{L}\p{N}]+)*/gu) || []).length;
}

function markDirty() {
  dirty.value = true;
  message.value = '';
  clearTimeout(wordTimer);
  wordTimer = setTimeout(() => {
    countWords();
    refreshChanges();
  }, 300);
}

function onInput(e) {
  restoreCommentMarkers(markersBefore, page.value);
  markersBefore = [];
  if (pendingJoin) {
    joinLeftParagraphs(pendingJoin);
    pendingJoin = null;
  }
  if (pendingSplit) {
    markSplit(pendingSplit); // Enter while tracking: the new paragraph break is an insertion
    pendingSplit = null;
  }
  // Never leave the page completely empty: keep at least one paragraph to type into.
  if (!page.value.firstElementChild) {
    const p = document.createElement('p');
    p.append(document.createElement('br'));
    page.value.append(p);
    placeCaret(p, true);
  }
  const range = selectionInPage();
  const control = range && controlAt(range.startContainer, page.value);
  if (control) clearPlaceholder(control, (id) => doc.isPlaceholderStyle(id));
  const cell = range && (range.startContainer.nodeType === 1 ? range.startContainer : range.startContainer.parentElement)?.closest('td.showing-placeholder');
  if (cell) {
    cell.classList.remove('showing-placeholder');
    for (const s of cell.querySelectorAll('[data-rstyle]')) if (doc.isPlaceholderStyle(s.dataset.rstyle)) delete s.dataset.rstyle;
  }
  doc.renumber(page.value);
  undoStack.schedule();
  markDirty();
  refreshComments();
  if (e?.inputType?.startsWith('format')) noteUntrackedFormatting();
}

function onBeforeInput(e) {
  // Undo/redo from the browser's Edit or context menu: use our history, not the native stack.
  if (e.inputType === 'historyUndo' || e.inputType === 'historyRedo') {
    e.preventDefault();
    if (e.inputType === 'historyUndo') undo();
    else redo();
    return;
  }
  if (!removesContent(e)) return;
  // Typing, deleting or pasting while the selection is in read-only content: say what it is.
  const locked = readOnlySelection();
  if (locked) {
    e.preventDefault();
    explainReadOnly(locked, true);
    return;
  }
  if (trackedBeforeInput(e)) return;
  const ranges = rangesRemovedBy(e);
  // Backspace or Delete next to hidden markers (comment ranges, bookmarks, field codes) would remove
  // just the invisible marker: delete what is beyond it instead.
  const sel = window.getSelection();
  const step = { deleteContentBackward: true, deleteContentForward: false }[e.inputType];
  if (step !== undefined && sel.isCollapsed && ranges.length === 1 && onlyHiddenMarkers(ranges[0])) {
    e.preventDefault();
    const target = beyondHiddenMarkers(ranges[0], step);
    if (!target || refuseRemoving([target])) return;
    markersBefore = commentMarkers(page.value);
    sel.removeAllRanges();
    sel.addRange(target);
    document.execCommand('delete'); // fires input, so onInput puts the markers back
    return;
  }
  // A deletion across paragraphs must join them; see joinLeftParagraphs.
  const first = ranges[0] && closestBlock(ranges[0].startContainer);
  const last = ranges[0] && closestBlock(ranges[0].endContainer);
  pendingJoin = first && last && first !== last ? { first, last, breaks: new Set(first.querySelectorAll('br')) } : null;
  // Typing over, deleting, cutting or dragging away a tracked change would drop it silently.
  if (refuseRemoving(ranges)) e.preventDefault();
  // Comment range ends deleted with the text are put back in onInput, at the caret: where the text
  // was deleted. (After a drag the caret is at the drop instead; the cleanup on save covers that.)
  else if (e.inputType !== 'deleteByDrag') markersBefore = commentMarkers(page.value);
}

let markersBefore = [];
let pendingJoin = null;

/**
 * The browser won't merge into a paragraph left holding only a hidden marker (a comment range or
 * bookmark start), so deleting from its start into the next paragraph left an invisible empty
 * paragraph behind. Join what the deletion left of the two, as the browser's own merge would.
 */
function joinLeftParagraphs({ first, last, breaks }) {
  if (!first.isConnected) return;
  let next = first.nextElementSibling;
  while (next?.classList.contains('raw-hidden')) next = next.nextElementSibling;
  if (last.isConnected && next === last && first.parentElement === last.parentElement) {
    const placeholder = (el) => el.lastChild?.tagName === 'BR' && !el.textContent;
    if (placeholder(first)) first.lastChild.remove();
    const kids = [...last.childNodes].filter((n) => !(n.tagName === 'BR' && n === last.lastChild && !last.textContent));
    first.append(...kids);
    last.remove();
  }
  // A browser fills a paragraph it has emptied with a placeholder <br> (Firefox keeps it when the
  // next paragraph joins on): one that now has text after it would be saved as a line break.
  for (const br of first.querySelectorAll('br')) {
    if (!breaks.has(br) && hasContentAfter(br, first)) br.remove();
  }
}

/** True if anything visible comes after `node` in `block`. */
function hasContentAfter(node, block) {
  const range = document.createRange();
  range.setStartAfter(node);
  range.setEnd(block, block.childNodes.length);
  const frag = range.cloneContents();
  return !!frag.textContent || !!frag.querySelector('br, img, .raw-inline:not(.raw-hidden), .sdt');
}

// ------------------------------------------------------------------ table rows

/** A new row at the end of the table (tracked while tracking), with the caret in its first cell. */
function addRow(cell) {
  const table = cell.closest('table');
  undoStack.flush();
  const { row, reason } = doc.addTableRow(table, trackedBy(cell));
  if (!row) {
    message.value = reason;
    return;
  }
  undoStack.commit();
  selectCell(row.querySelector('td[contenteditable="true"]') || row.cells[0]);
  markDirty();
}

// ------------------------------------------------------------------ picture layout

const HANDLES = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'];
const CM = 360000; // EMU per centimetre
const PX = 9525; // EMU per pixel
const picture = reactive({ el: null, x: 0, y: 0, w: 0, h: 0, wrap: 'inline', cropping: false });
const picturePanel = reactive({
  open: false, x: 0, y: 0, wrap: 'inline', hAlign: 'left', hOffset: 0, vAlign: 'offset', vOffset: 0, width: 0, height: 0, lock: true, alt: '',
  cropL: 0, cropT: 0, cropR: 0, cropB: 0, rot: 0, flipH: false, flipV: false,
  border: 'none', borderColor: '#000000', shadow: false, glow: false, softEdge: false, reflection: false, colour: 'normal', brightness: 0, contrast: 0,
});
const BORDERS = { thin: 0.75, medium: 1.5, thick: 3 };
const picturePanelEl = ref(null);
const layoutOptionsButton = ref(null);
const round = (n, d = 2) => Math.round(n * 10 ** d) / 10 ** d;

/** The picture a selection is on: exactly one picture and nothing else (as a click selects it). */
function pictureInSelection(range) {
  if (!range || range.collapsed) return null;
  const frag = range.cloneContents();
  if (frag.textContent.trim() || frag.querySelectorAll('.picture').length !== 1) return null;
  return [...page.value.querySelectorAll('.picture')].find((p) => range.intersectsNode(p)) || null;
}

/** Selects a picture, as a click does in Word: the frame with its handles shows around it. */
function selectPicture(span) {
  const range = document.createRange();
  range.selectNode(span);
  page.value.focus({ preventScroll: true });
  const sel = window.getSelection();
  sel.removeAllRanges();
  sel.addRange(range);
  savedRange = range.cloneRange();
  showPicture(span);
}

function showPicture(span) {
  if (picture.el !== span) {
    picturePanel.open = false;
    picture.cropping = false;
  }
  picture.el = span;
  updatePictureFrame();
  if (span) message.value = 'Picture: drag a corner to resize, or open Layout Options for wrapping and position';
}

/** Puts the frame over the selected picture (in the sheet's coordinates). */
function updatePictureFrame() {
  const span = picture.el;
  if (!span?.isConnected) {
    picture.el = null;
    picturePanel.open = false;
    return;
  }
  const sheet = page.value.closest('.docx-sheet').getBoundingClientRect();
  const r = span.getBoundingClientRect(); // the picture's frame (what shows of it, cropped)
  Object.assign(picture, { x: r.left - sheet.left, y: r.top - sheet.top, w: r.width, h: r.height, wrap: span.dataset.wrap || 'inline' });
}

/** Applies a change to the selected picture as one undoable step, keeping it selected. */
function changePicture(change, done) {
  const span = picture.el;
  if (!span?.isConnected) return;
  undoStack.flush();
  change(span);
  undoStack.commit();
  markDirty();
  nextTick(() => {
    if (span.isConnected) selectPicture(span);
    fillPicturePanel();
    if (done) message.value = done;
  });
}

/** Drag a handle: corners keep the proportions, sides stretch, as in Word. */
function startResize(e, handle) {
  const span = picture.el;
  const info = span && doc.pictureLayout(span);
  if (!info) return;
  const look = doc.pictureLookOf(span);
  const start = { x: e.clientX, y: e.clientY, w: info.cx / PX, h: info.cy / PX };
  const target = e.currentTarget;
  target.setPointerCapture(e.pointerId);
  let size = { w: start.w, h: start.h };
  const move = (ev) => {
    const dx = (ev.clientX - start.x) * (handle.includes('w') ? -1 : handle.includes('e') ? 1 : 0);
    const dy = (ev.clientY - start.y) * (handle.includes('n') ? -1 : handle.includes('s') ? 1 : 0);
    let w = start.w + dx;
    let h = start.h + dy;
    if (handle.length === 2) {
      const scale = Math.max(w / start.w, h / start.h);
      w = start.w * scale;
      h = start.h * scale;
    }
    size = { w: Math.max(8, w), h: Math.max(8, h) };
    doc.stylePicture(span, { ...info, cx: size.w * PX, cy: size.h * PX }, look); // a preview until it is let go
    updatePictureFrame();
  };
  const up = () => {
    target.removeEventListener('pointermove', move);
    target.removeEventListener('pointerup', up);
    target.removeEventListener('pointercancel', up);
    if (size.w === start.w && size.h === start.h) return;
    changePicture((s) => doc.setPictureSize(s, size.w * PX, size.h * PX), `Picture resized to ${round(size.w * PX / CM, 1)} × ${round(size.h * PX / CM, 1)} cm`);
  };
  target.addEventListener('pointermove', move);
  target.addEventListener('pointerup', up);
  target.addEventListener('pointercancel', up);
}

/**
 * A press on a picture selects it; dragging a selected floating picture moves it, its new place
 * measured from the left margin and from its paragraph (so it moves with the text, as by default
 * in Word).
 */
function onPictureMousedown(e, span) {
  e.preventDefault();
  const wasSelected = picture.el === span;
  selectPicture(span);
  if (!wasSelected || (span.dataset.wrap || 'inline') === 'inline' || e.button !== 0) return;
  const start = { x: e.clientX, y: e.clientY };
  let moved = false;
  const move = (ev) => {
    const dx = ev.clientX - start.x;
    const dy = ev.clientY - start.y;
    if (!moved && Math.hypot(dx, dy) < 4) return;
    moved = true;
    span.style.transform = `translate(${dx}px, ${dy}px) ${span.dataset.turn || ''}`;
    updatePictureFrame();
  };
  const up = (ev) => {
    window.removeEventListener('pointermove', move);
    window.removeEventListener('pointerup', up);
    if (!moved) return;
    const r = span.getBoundingClientRect(); // where it was dropped
    span.style.transform = span.dataset.turn || '';
    const sheet = page.value.closest('.docx-sheet').getBoundingClientRect();
    const block = span.closest(BLOCK_SELECTOR) || span.parentElement;
    const left = r.left - (sheet.left + doc.page.left);
    const top = r.top - block.getBoundingClientRect().top;
    changePicture((s) => doc.setPicturePosition(s, {
      h: { relativeFrom: 'margin', offset: Math.round(left) * PX },
      v: { relativeFrom: 'paragraph', offset: Math.round(top) * PX },
    }), 'Picture moved');
    void ev;
  };
  window.addEventListener('pointermove', move);
  window.addEventListener('pointerup', up);
}

/** Fills the Layout Options panel from the selected picture. */
function fillPicturePanel() {
  if (!picture.el) return;
  const info = doc.pictureLayout(picture.el);
  if (!info) return;
  const page = doc.page;
  const h = info.h || { relativeFrom: 'margin', align: 'left' };
  const v = info.v || { relativeFrom: 'paragraph', offset: 0 };
  let hAlign = 'word';
  let hOffset = 0;
  if (h.align && ['left', 'center', 'right'].includes(h.align) && ['margin', 'column'].includes(h.relativeFrom)) hAlign = h.align;
  else if (h.align == null && ['margin', 'column', 'page'].includes(h.relativeFrom)) {
    hAlign = 'offset';
    hOffset = round(((h.offset || 0) / PX - (h.relativeFrom === 'page' ? page.left : 0)) * PX / CM);
  }
  let vAlign = 'word';
  if (v.align && ['top', 'center', 'bottom'].includes(v.align) && v.relativeFrom === 'margin') vAlign = v.align;
  else if (v.align == null && ['paragraph', 'line'].includes(v.relativeFrom)) vAlign = 'offset';
  const look = doc.pictureLookOf(picture.el);
  const border = look.border ? (look.border.width <= 1 ? 'thin' : look.border.width <= 2 ? 'medium' : 'thick') : 'none';
  Object.assign(picturePanel, {
    wrap: info.wrap, hAlign, hOffset, vAlign, vOffset: vAlign === 'offset' ? round((v.offset || 0) / CM) : 0,
    width: round(info.cx / CM), height: round(info.cy / CM), alt: info.descr,
    cropL: round(look.crop.l * 100, 1), cropT: round(look.crop.t * 100, 1), cropR: round(look.crop.r * 100, 1), cropB: round(look.crop.b * 100, 1),
    rot: round(look.rot, 1), flipH: look.flipH, flipV: look.flipV,
    border, borderColor: `#${look.border?.color || '000000'}`.toLowerCase(), shadow: look.shadow, glow: !!look.glow, softEdge: look.softEdge,
    reflection: look.reflection, colour: look.greyscale ? 'greyscale' : 'normal', brightness: round(look.brightness * 100), contrast: round(look.contrast * 100),
  });
}

function togglePicturePanel() {
  if (picturePanel.open) return closePicturePanel();
  fillPicturePanel();
  const r = layoutOptionsButton.value.getBoundingClientRect();
  Object.assign(picturePanel, { open: true, x: Math.max(8, Math.min(r.right + 6, window.innerWidth - 310)), y: Math.max(8, r.top) });
  nextTick(() => {
    const el = picturePanelEl.value;
    if (!el) return;
    picturePanel.y = Math.max(8, Math.min(picturePanel.y, window.innerHeight - el.offsetHeight - 8)); // all of it on screen
    el.querySelector('input:checked, input')?.focus();
  });
}

function closePicturePanel() {
  picturePanel.open = false;
  if (picture.el?.isConnected) selectPicture(picture.el);
  else page.value.focus();
}

const WRAP_NAMES = Object.fromEntries(WRAPS);

function applyWrap(wrap) {
  // Tight and Through: the text wraps to the picture's outline (what isn't transparent), as in Word.
  const outline = (wrap === 'tight' || wrap === 'through') && picture.el
    ? pictureOutline(picture.el.querySelector('img'), doc.pictureLookOf(picture.el).crop) : null;
  changePicture((s) => doc.setPictureWrap(s, wrap, outline), `Text wrapping: ${WRAP_NAMES[wrap]}${outline ? ', around its outline' : ''}`);
}

function applyPosition(which) {
  const p = picturePanel;
  const h = which === 'h' && p.hAlign !== 'word'
    ? (p.hAlign === 'offset' ? { relativeFrom: 'margin', offset: (Number(p.hOffset) || 0) * CM } : { relativeFrom: 'margin', align: p.hAlign })
    : null;
  const v = which === 'v' && p.vAlign !== 'word'
    ? (p.vAlign === 'offset' ? { relativeFrom: 'paragraph', offset: (Number(p.vOffset) || 0) * CM } : { relativeFrom: 'margin', align: p.vAlign })
    : null;
  if (!h && !v) return;
  changePicture((s) => doc.setPicturePosition(s, { h, v }), 'Picture positioned');
}

function applySize(which) {
  const info = doc.pictureLayout(picture.el);
  if (!info) return;
  let w = Math.max(0.1, Number(picturePanel.width) || 0.1) * CM;
  let h = Math.max(0.1, Number(picturePanel.height) || 0.1) * CM;
  if (picturePanel.lock) {
    if (which === 'width') h = (w * info.cy) / info.cx;
    else w = (h * info.cx) / info.cy;
  }
  changePicture((s) => doc.setPictureSize(s, w, h), `Picture resized to ${round(w / CM, 1)} × ${round(h / CM, 1)} cm`);
}

/** Back to the picture's own size (at 96 pixels an inch), narrowed to fit between the margins. */
function resetPictureSize() {
  const img = picture.el?.querySelector('img');
  if (!img?.naturalWidth) return;
  const max = doc.page.width - doc.page.left - doc.page.right;
  const scale = Math.min(1, max / img.naturalWidth);
  changePicture((s) => doc.setPictureSize(s, img.naturalWidth * scale * PX, img.naturalHeight * scale * PX), 'Picture back to its own size');
}

/** Crop with the handles, as Word's Crop button: dragging a handle cuts (or uncuts) that side. */
function toggleCropping() {
  picture.cropping = !picture.cropping;
  message.value = picture.cropping ? 'Crop: drag the handles to cut the picture; Esc or Done cropping when finished' : 'Picture cropped';
}

function startCrop(e, handle) {
  const span = picture.el;
  const info = span && doc.pictureLayout(span);
  if (!info) return;
  const look = doc.pictureLookOf(span);
  const c0 = look.crop;
  const fullW = info.cx / PX / (1 - c0.l - c0.r);
  const fullH = info.cy / PX / (1 - c0.t - c0.b);
  const start = { x: e.clientX, y: e.clientY };
  const target = e.currentTarget;
  target.setPointerCapture(e.pointerId);
  let crop = c0;
  const move = (ev) => {
    const dx = (ev.clientX - start.x) / fullW;
    const dy = (ev.clientY - start.y) / fullH;
    crop = {
      l: handle.includes('w') ? c0.l + dx : c0.l,
      r: handle.includes('e') ? c0.r - dx : c0.r,
      t: handle.includes('n') ? c0.t + dy : c0.t,
      b: handle.includes('s') ? c0.b - dy : c0.b,
    };
    for (const side of ['l', 'r', 't', 'b']) crop[side] = Math.max(-0.5, Math.min(0.9, crop[side]));
    if (crop.l + crop.r > 0.95 || crop.t + crop.b > 0.95) return;
    const w = fullW * (1 - crop.l - crop.r);
    const h = fullH * (1 - crop.t - crop.b);
    doc.stylePicture(span, { ...info, cx: w * PX, cy: h * PX }, { ...look, crop });
    updatePictureFrame();
  };
  const up = () => {
    target.removeEventListener('pointermove', move);
    target.removeEventListener('pointerup', up);
    if (crop === c0) return;
    changePicture((s) => doc.setPictureCrop(s, crop), 'Picture cropped');
    nextTick(() => { picture.cropping = true; });
  };
  target.addEventListener('pointermove', move);
  target.addEventListener('pointerup', up);
}

/** Drag the round handle above a picture to rotate it (Shift: in steps of 15°), as in Word. */
function startRotate(e) {
  const span = picture.el;
  const info = span && doc.pictureLayout(span);
  if (!info) return;
  const look = doc.pictureLookOf(span);
  const r = span.getBoundingClientRect();
  const center = { x: r.left + r.width / 2, y: r.top + r.height / 2 };
  const target = e.currentTarget;
  target.setPointerCapture(e.pointerId);
  let rot = look.rot;
  const move = (ev) => {
    let deg = (Math.atan2(ev.clientY - center.y, ev.clientX - center.x) * 180) / Math.PI + 90;
    if (ev.shiftKey) deg = Math.round(deg / 15) * 15;
    rot = ((Math.round(deg) % 360) + 360) % 360;
    doc.stylePicture(span, info, { ...look, rot });
    updatePictureFrame();
  };
  const up = () => {
    target.removeEventListener('pointermove', move);
    target.removeEventListener('pointerup', up);
    if (rot === look.rot) return;
    changePicture((s) => doc.setPictureRotation(s, { rot, flipH: look.flipH, flipV: look.flipV }), `Picture rotated to ${rot}°`);
  };
  target.addEventListener('pointermove', move);
  target.addEventListener('pointerup', up);
}

function applyCrop() {
  const p = picturePanel;
  const crop = { l: (Number(p.cropL) || 0) / 100, t: (Number(p.cropT) || 0) / 100, r: (Number(p.cropR) || 0) / 100, b: (Number(p.cropB) || 0) / 100 };
  changePicture((s) => doc.setPictureCrop(s, crop), 'Picture cropped');
}

function applyRotation() {
  const p = picturePanel;
  changePicture((s) => doc.setPictureRotation(s, { rot: Number(p.rot) || 0, flipH: p.flipH, flipV: p.flipV }), `Picture rotated to ${((Math.round(Number(p.rot) || 0) % 360) + 360) % 360}°`);
}

function rotateBy(deg) {
  picturePanel.rot = (((Number(picturePanel.rot) || 0) + deg) % 360 + 360) % 360;
  applyRotation();
}

function flip(axis) {
  picturePanel[axis] = !picturePanel[axis];
  const p = picturePanel;
  changePicture((s) => doc.setPictureRotation(s, { rot: Number(p.rot) || 0, flipH: p.flipH, flipV: p.flipV }), axis === 'flipH' ? 'Picture flipped horizontally' : 'Picture flipped vertically');
}

function applyEffects() {
  const p = picturePanel;
  const look = {
    border: p.border === 'none' ? null : { color: p.borderColor.replace('#', ''), width: BORDERS[p.border] },
    shadow: p.shadow, glow: p.glow ? 'FFC000' : null, softEdge: p.softEdge, reflection: p.reflection,
    greyscale: p.colour === 'greyscale', brightness: (Number(p.brightness) || 0) / 100, contrast: (Number(p.contrast) || 0) / 100,
  };
  changePicture((s) => doc.setPictureEffects(s, look), 'Picture effects changed');
}

function applyAlt() {
  changePicture((s) => doc.setPictureDescription(s, picturePanel.alt.trim()), picturePanel.alt.trim() ? 'Alt text saved' : 'Alt text removed');
}

// ------------------------------------------------------------------ table layout

const tableMenu = reactive({ open: false, x: 0, y: 0, rows: 1, cols: 1, newCols: 1, multi: false });
const tableMenuEl = ref(null);
const layoutButton = ref(null);

const cellOf = (node) => (node?.nodeType === 1 ? node : node?.parentElement)?.closest('td') || null;

/**
 * The table cells selected (or the one with the caret): the table, the cell the selection starts
 * in, and the rows and grid columns covered ([first, last]). Null outside a table.
 */
function tableSelection() {
  const range = selectionInPage() || savedRange;
  const a = range && cellOf(range.startContainer);
  if (!a || !page.value.contains(a)) return null;
  const table = a.closest('table');
  if (table.dataset.raw == null) return null;
  let b = cellOf(range.endContainer);
  // A selection ending at the very start of a cell (as when Tab selects a whole cell) ends before it.
  const end = range.cloneRange();
  end.collapse(false);
  if (b && b !== a && isAtBlockStart(end, b) && !isEmptyBlock(b)) {
    const cells = [...table.querySelectorAll(':scope > tbody > tr > td, :scope > tr > td')];
    b = cells[cells.indexOf(b) - 1] || a;
  }
  if (!b || b.closest('table') !== table) b = a;
  const pa = doc.cellPosition(a);
  const pb = doc.cellPosition(b);
  if (!pa || !pb) return { table, cell: a, rows: [0, 0], cols: [0, 0], newCols: 1, multi: false };
  const cols = [Math.min(pa.start, pb.start), Math.max(pa.start + pa.span, pb.start + pb.span) - 1];
  return {
    table,
    cell: a,
    newCols: a === b ? 1 : cols[1] - cols[0] + 1, // columns to insert: one for a single (even spanning) cell
    multi: a !== b,
    rows: [Math.min(pa.row, pb.row), Math.max(pa.row, pb.row)],
    cols: [Math.min(pa.start, pb.start), Math.max(pa.start + pa.span, pb.start + pb.span) - 1],
  };
}

function toggleTableMenu() {
  if (tableMenu.open) return closeTableMenu();
  const sel = tableSelection();
  if (!sel) return;
  const rect = layoutButton.value.getBoundingClientRect();
  Object.assign(tableMenu, {
    open: true,
    x: Math.max(8, Math.min(rect.left, window.innerWidth - 240)),
    y: rect.bottom + 4,
    rows: sel.rows[1] - sel.rows[0] + 1,
    cols: sel.cols[1] - sel.cols[0] + 1,
    newCols: sel.newCols,
    multi: sel.multi,
  });
  nextTick(() => tableMenuEl.value?.querySelector('button')?.focus());
}

function closeTableMenu() {
  tableMenu.open = false;
  page.value.focus();
  restoreSelection();
}

/** Up and Down (Home, End) move through a menu's items. */
function onMenuArrows(e) {
  if (!['ArrowDown', 'ArrowUp', 'Home', 'End'].includes(e.key)) return;
  e.preventDefault();
  const items = [...e.currentTarget.querySelectorAll('button')];
  const i = items.indexOf(document.activeElement);
  const next = { ArrowDown: (i + 1) % items.length, ArrowUp: (i - 1 + items.length) % items.length, Home: 0, End: items.length - 1 }[e.key];
  items[next].focus();
}

/** The rendered cell of `row` that covers grid column `col` (or its last cell). */
function cellAtColumn(row, col) {
  const cells = [...row.cells];
  return cells.find((td) => {
    const p = doc.cellPosition(td);
    return p && p.start <= col && col < p.start + p.span;
  }) || cells[cells.length - 1] || null;
}

const plural = (n, one, many = `${one}s`) => (n === 1 ? one : `${n} ${many}`);

/** Insert or delete rows, columns or the table, from the Table layout menu (as Word's Layout tab). */
function tableAction(kind) {
  tableMenu.open = false;
  refocus();
  const sel = tableSelection();
  if (!sel) {
    message.value = 'Click in a table first';
    return;
  }
  const locked = readOnlyAt(sel.cell, page.value);
  if (locked) return explainReadOnly(locked, true);
  const { table, cell } = sel;
  const [r0, r1] = sel.rows;
  const [c0, c1] = sel.cols;
  const rowCount = r1 - r0 + 1;
  const colCount = c1 - c0 + 1;
  const tracked = trackedBy(cell);
  const neighbours = { before: table.previousElementSibling, after: table.nextElementSibling, parent: table.parentElement };
  const row = cell.parentElement;
  undoStack.flush();
  let res;
  let focus = null;
  let done = '';
  if (kind === 'rowsAbove' || kind === 'rowsBelow') {
    const below = kind === 'rowsBelow';
    res = doc.insertTableRows(table, below ? r1 : r0, rowCount, below, tracked);
    focus = res.rows && cellAtColumn(res.rows[0], c0);
    done = `${rowCount === 1 ? 'Row' : `${rowCount} rows`} inserted ${below ? 'below' : 'above'}`;
  } else if (kind === 'columnsLeft' || kind === 'columnsRight') {
    const right = kind === 'columnsRight';
    res = doc.insertTableColumns(table, c0, colCount, sel.newCols, right, tracked);
    if (!res.reason) focus = cellAtColumn(row, right ? c1 + 1 : c0);
    done = `${sel.newCols === 1 ? 'Column' : `${sel.newCols} columns`} inserted to the ${right ? 'right' : 'left'}${tracked ? ' (tracked)' : ''}`;
  } else if (kind === 'deleteRows') {
    res = doc.deleteTableRows(table, r0, r1, tracked);
    done = tracked ? `${plural(rowCount, 'Row', 'rows')} marked as deleted (tracked)` : `${plural(rowCount, 'Row', 'rows')} deleted – Undo (Ctrl+Z) brings ${rowCount === 1 ? 'it' : 'them'} back`;
    if (!res.reason && !res.removed) {
      const rows = doc.tableRows(table);
      focus = row.isConnected ? cellAtColumn(row, c0) : rows[Math.min(r0, rows.length - 1)] && cellAtColumn(rows[Math.min(r0, rows.length - 1)], c0);
    }
  } else if (kind === 'deleteColumns') {
    res = doc.deleteTableColumns(table, c0, c1, tracked);
    done = tracked && !res.removed && table.querySelector('td[data-rev-cell="del"]')
      ? `${plural(colCount, 'Column', 'columns')} marked as deleted (tracked)`
      : `${plural(colCount, 'Column', 'columns')} deleted – Undo (Ctrl+Z) brings ${colCount === 1 ? 'it' : 'them'} back`;
    if (!res.reason && !res.removed && row.isConnected) focus = cellAtColumn(row, Math.max(0, c0 - 1));
  } else if (kind === 'merge') {
    res = doc.mergeTableCells(table, sel.rows, sel.cols, tracked);
    if (res.cell) focus = res.cell;
    // Word documents can record cells merged down a column as a tracked change, but not across a row.
    done = !tracked ? 'Cells merged'
      : res.down && !res.across ? 'Cells merged (tracked)'
        : res.down ? 'Cells merged – tracked down the rows; merging across a row can’t be tracked in a Word document'
          : 'Cells merged – merging across a row can’t be tracked in a Word document, so it isn’t marked';
  } else if (kind === 'split') {
    const { cols, rows } = splitDialog;
    res = doc.splitTableCell(table, cell, cols, rows, tracked);
    focus = cell.isConnected ? cell : null;
    done = `Cell split into ${cols} × ${rows}${tracked ? ' (tracked)' : ''}`;
  } else if (kind === 'deleteTable') {
    res = doc.deleteTable(table, tracked);
    done = tracked ? 'Table marked as deleted (tracked)' : 'Table deleted – Undo (Ctrl+Z) brings it back';
  }
  if (res.reason) {
    message.value = res.reason;
    return;
  }
  if (res.removed) {
    // The caret goes where the table was: the paragraph after it (or before it).
    if (!page.value.firstElementChild) page.value.append(doc.emptyParagraph());
    const next = [neighbours.after, neighbours.before].find((n) => n?.isConnected);
    const target = next?.matches(BLOCK_SELECTOR) ? next : next?.querySelector(BLOCK_SELECTOR) || neighbours.parent.querySelector(BLOCK_SELECTOR) || page.value.querySelector(BLOCK_SELECTOR);
    if (target) placeCaret(target, next !== neighbours.before);
  } else if (focus) {
    placeCaret(focus.querySelector(BLOCK_SELECTOR) || focus, true);
  } else if (cell.isConnected) {
    placeCaret(cell.querySelector(BLOCK_SELECTOR) || cell, true);
  }
  doc.renumber(page.value);
  undoStack.commit();
  markDirty();
  refreshChanges();
  refreshComments();
  onSelectionChange();
  message.value = done;
}

const splitDialog = reactive({ open: false, x: 0, y: 0, cols: 2, rows: 1 });
const splitDialogEl = ref(null);

/** Split cells…: asks how many columns and rows, as Word does. */
function openSplitDialog() {
  tableMenu.open = false;
  const rect = layoutButton.value.getBoundingClientRect();
  Object.assign(splitDialog, { open: true, cols: 2, rows: 1, x: Math.max(8, Math.min(rect.left, window.innerWidth - 240)), y: rect.bottom + 4 });
  nextTick(() => splitDialogEl.value?.querySelector('input')?.select());
}

function closeSplitDialog() {
  splitDialog.open = false;
  refocus();
}

function splitCells() {
  splitDialog.cols = Math.min(20, Math.max(1, Math.floor(Number(splitDialog.cols) || 1)));
  splitDialog.rows = Math.min(50, Math.max(1, Math.floor(Number(splitDialog.rows) || 1)));
  splitDialog.open = false;
  if (splitDialog.cols === 1 && splitDialog.rows === 1) return refocus();
  tableAction('split');
}

// ------------------------------------------------------------------ inserting tables

const PICKER_ROWS = 8;
const PICKER_COLS = 10;
const MAX_COLS = 63; // as many as Word allows
const MAX_ROWS = 500;
const tablePicker = reactive({ open: false, x: 0, y: 0, rows: 1, cols: 1, customRows: 2, customCols: 3 });
const tablePickerEl = ref(null);
const tableButton = ref(null);

function toggleTablePicker() {
  if (tablePicker.open) return closeTablePicker();
  const rect = tableButton.value.getBoundingClientRect();
  Object.assign(tablePicker, { open: true, rows: 1, cols: 1, x: Math.max(8, Math.min(rect.left - 80, window.innerWidth - 250)), y: rect.bottom + 4 });
  nextTick(() => tablePickerEl.value?.querySelector('.cell')?.focus());
}

function closeTablePicker() {
  tablePicker.open = false;
  page.value.focus();
  restoreSelection();
}

/** Arrow keys grow or shrink the table in the picker, as in Word; Enter inserts it. */
function onPickerKeydown(e) {
  const step = { ArrowRight: [0, 1], ArrowLeft: [0, -1], ArrowDown: [1, 0], ArrowUp: [-1, 0] }[e.key];
  if (!step) return;
  e.preventDefault();
  tablePicker.rows = Math.min(PICKER_ROWS, Math.max(1, tablePicker.rows + step[0]));
  tablePicker.cols = Math.min(PICKER_COLS, Math.max(1, tablePicker.cols + step[1]));
  nextTick(() => tablePickerEl.value?.querySelectorAll('.cell')[(tablePicker.rows - 1) * PICKER_COLS + tablePicker.cols - 1]?.focus());
}

/**
 * Inserts an empty table of `rows` × `cols` at the caret, as Word does: before the paragraph if the
 * caret is at its start (or it is empty), otherwise the paragraph is split there and the table goes
 * between. While tracking, its rows are tracked insertions. The caret goes into its first cell.
 */
function insertTable(rows, cols) {
  tablePicker.open = false;
  refocus();
  rows = Math.min(MAX_ROWS, Math.max(1, Math.floor(Number(rows) || 1)));
  cols = Math.min(MAX_COLS, Math.max(1, Math.floor(Number(cols) || 1)));
  const range = selectionInPage();
  if (!range) {
    message.value = 'Click in the text where the table should go';
    return;
  }
  const locked = readOnlySelection();
  if (locked) return explainReadOnly(locked, true);
  if (controlAt(range.startContainer, page.value)) {
    message.value = 'A table can’t go in a field';
    return;
  }
  const tracked = trackedBy(range.startContainer);
  undoStack.flush();
  range.collapse(true);
  let block = closestBlock(range.startContainer);
  if (!block) {
    block = doc.emptyParagraph();
    page.value.append(block);
  } else if (!isEmptyBlock(block) && !isAtBlockStart(range, block)) {
    document.execCommand('insertParagraph');
    if (tracked) markSplit(tracked);
    block = closestBlock(selectionInPage().startContainer) || block;
  }
  const table = doc.newTable(rows, cols, tracked);
  block.before(table);
  refreshCss(); // the Table Grid style may be new
  placeCaret(table.rows[0].cells[0].querySelector(BLOCK_SELECTOR) || table.rows[0].cells[0], true);
  doc.renumber(page.value);
  undoStack.commit();
  markDirty();
  refreshChanges();
  onSelectionChange();
  message.value = `Table inserted: ${cols} × ${rows}${tracked ? ' (tracked)' : ''}`;
}

// ------------------------------------------------------------------ inserting pictures

const pictureInput = ref(null);

function pickPicture() {
  pictureInput.value.value = '';
  pictureInput.value.click();
}

function onPicturesPicked(e) {
  const files = [...e.target.files];
  e.target.value = '';
  if (files.length) insertPictures(files);
}

/**
 * Inserts pictures at the caret (or `at`, a range), in line with the text and replacing what is
 * selected, as Word does; tracked while tracking.
 */
async function insertPictures(files, at = null) {
  if (at) {
    page.value.focus();
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(at);
  } else {
    refocus();
  }
  const range = selectionInPage();
  if (!range) {
    message.value = 'Click in the text where the picture should go';
    return;
  }
  const locked = readOnlySelection();
  if (locked) return explainReadOnly(locked, true);
  if (controlAt(range.startContainer, page.value)) {
    message.value = 'A picture can’t go in a field: insert it in Word';
    return;
  }
  if (!range.collapsed && refuseRemoving([range])) return;
  const where = range.cloneRange();
  const made = [];
  const reasons = [];
  for (const file of files) {
    const res = await doc.addPicture(file);
    if (res.span) made.push(res.span);
    else reasons.push(res.reason);
  }
  if (!made.length) {
    message.value = reasons[0] || 'No picture was inserted';
    return;
  }
  page.value.focus();
  const sel = window.getSelection();
  sel.removeAllRanges();
  sel.addRange(where);
  undoStack.flush();
  const tracked = trackedBy(where.startContainer);
  if (!where.collapsed) {
    if (tracked) deleteTracked(where, false, tracked);
    else deleteSelection();
  }
  if (tracked) insertTracked(made, tracked);
  else insertAtCaret(made);
  page.value.normalize();
  undoStack.commit();
  markDirty();
  refreshChanges();
  onSelectionChange();
  message.value = `${made.length === 1 ? 'Picture' : `${made.length} pictures`} inserted${tracked ? ' (tracked)' : ''}${reasons.length ? ` – ${reasons[0]}` : ''}`;
}

// ------------------------------------------------------------------ adding comments

const composer = reactive({ open: false });
let composerRange = null; // where the new comment goes

/** With nothing selected, a comment goes on the word at the caret, as in Word. */
function wordAround(range) {
  const { startContainer: node, startOffset: offset } = range;
  if (!range.collapsed || node.nodeType !== 3) return range;
  const letter = /[\p{L}\p{N}_'’-]/u;
  let s = offset;
  let e = offset;
  while (s > 0 && letter.test(node.data[s - 1])) s--;
  while (e < node.length && letter.test(node.data[e])) e++;
  if (s === e) return range;
  const word = document.createRange();
  word.setStart(node, s);
  word.setEnd(node, e);
  return word;
}

/** New comment (Ctrl+Alt+M): opens a box for it in the comments panel, on the selected text. */
function startComment() {
  restoreSelection();
  const range = selectionInPage();
  if (!range) {
    message.value = 'Click in the text, or select some, to comment on it';
    return;
  }
  const locked = readOnlySelection();
  if (locked) return explainReadOnly(locked, true);
  const hostOf = (n) => (n.nodeType === 1 ? n : n.parentElement).closest('[contenteditable="true"]');
  if (hostOf(range.startContainer) !== hostOf(range.endContainer)) {
    message.value = 'A comment can’t run from a table cell into the text around it: select within one';
    return;
  }
  composerRange = wordAround(range).cloneRange();
  composer.open = true;
  commentsOpen.value = true;
  paintComments();
}

function postComment({ text, name }) {
  if (name) {
    author.value = name;
    named.value = true;
    try { localStorage.setItem(AUTHOR_KEY, name); } catch { /* this session only */ }
  }
  undoStack.flush();
  const comment = doc.addComment(composerRange, { author: author.value, text, date: revisionDate() });
  composer.open = false;
  composerRange = null;
  undoStack.commit();
  markDirty();
  refreshComments();
  selectComment(comment.id, false);
  page.value.focus();
  message.value = 'Comment added';
}

function cancelComment() {
  if (!composer.open) return;
  composer.open = false;
  composerRange = null;
  paintComments();
  page.value.focus();
  restoreSelection();
}

/** Runs a change to comments as one undoable step, then refreshes the panel. */
function changeComments(change, done) {
  undoStack.flush();
  change();
  undoStack.commit();
  markDirty();
  refreshComments();
  message.value = done;
}

/** A reply to thread `id`, as in Word (written as a real reply, linked to its comment). */
function replyComment({ id, text, name }) {
  if (name) {
    author.value = name;
    named.value = true;
    try { localStorage.setItem(AUTHOR_KEY, name); } catch { /* this session only */ }
  }
  let reply = null;
  changeComments(() => { reply = doc.addReply(id, { author: author.value, text, date: revisionDate() }, page.value); }, 'Reply added');
  if (!reply) message.value = 'This comment isn’t in the text here: reply to it in Word';
}

/** A comment's new text, kept on its mark until saved, so Undo covers it. */
function editComment({ id, text }) {
  const mark = commentMark(id);
  if (!mark || !text) return;
  changeComments(() => { mark.dataset.edit = text; }, 'Comment edited');
}

/** Resolves or reopens a thread, kept on its mark until saved, so Undo covers it. */
function resolveComment({ id, done }) {
  const mark = commentMark(id);
  if (!mark) return;
  changeComments(() => { mark.dataset.done = done ? '1' : '0'; }, done ? 'Comment resolved' : 'Comment reopened');
}

/** Deletes a comment and its replies (their marks go from the text); Undo brings them back. */
function deleteComment(id) {
  undoStack.flush();
  for (const el of page.value.querySelectorAll(`[data-comment-start="${CSS.escape(id)}"], [data-comment-end="${CSS.escape(id)}"], [data-comment-ref="${CSS.escape(id)}"]`)) el.remove();
  undoStack.commit();
  markDirty();
  refreshComments();
  message.value = 'Comment deleted – Undo (Ctrl+Z) brings it back';
}

// ------------------------------------------------------------------ accepting and rejecting changes

const changes = reactive({ total: 0, wordOnly: 0 });
const resolved = reactive({ accepted: 0, rejected: 0 }); // this session, for the review
const changeMenu = reactive({ open: null, x: 0, y: 0 });
const changeMenuEl = ref(null);
const acceptMenuButton = ref(null);
const rejectMenuButton = ref(null);
const changesTitle = computed(() => {
  const parts = [`${changes.total} tracked ${changes.total === 1 ? 'change' : 'changes'} to accept or reject here`];
  if (changes.wordOnly) parts.push(`${changes.wordOnly} more (in locked content) only in Word`);
  return parts.join('; ');
});

function refreshChanges() {
  if (!page.value) return;
  changes.total = changeItems(page.value).length;
  changes.wordOnly = doc.wordOnlyChanges(page.value);
}

const revisionOfItem = (item) => doc.revisionOf(item.el, ['run', 'para', 'row', 'cells', 'merge', 'tprops', 'rprops', 'cprops'].includes(item.kind) ? item.kind : null);

/** "the insertion by Alice", "your deletion", "the formatting change by Eve", "the inserted row by Dan". */
function describeChange(item) {
  const rev = revisionOfItem(item);
  if (!rev) return 'the change';
  const who = rev.own ? 'your' : 'the';
  const by = rev.own || !rev.author ? '' : ` by ${rev.author}`;
  if (item.kind === 'run') return `the formatting change${by}`;
  if (item.kind === 'para') return `the paragraph formatting change${by}`;
  if (item.kind === 'row') return `the ${rev.kind === 'ins' ? 'inserted' : 'deleted'} row${by}`;
  if (item.kind === 'cells') return `the ${rev.kind === 'ins' ? 'inserted' : 'deleted'} column${by}`;
  if (item.kind === 'merge') return `the ${item.el.dataset.revMerge === 'merge' || mergeGroup(item.el).some((c) => c.dataset.revMerge === 'merge') ? 'merged' : 'split'} cells${by}`;
  if (item.kind === 'tprops') return `the table formatting change${by}`;
  if (item.kind === 'rprops') return `the row formatting change${by}`;
  if (item.kind === 'cprops') return `the cell formatting change${by}`;
  if (rev.mark) return `${who} paragraph break ${rev.kind === 'ins' ? 'insertion' : 'deletion'}${by}`;
  return `${who} ${rev.kind === 'ins' ? 'insertion' : 'deletion'}${by}`;
}

/** A caret left between paragraphs (or before a table) moves into the text next to it. */
function settleCaret(range) {
  const { startContainer: node, startOffset: offset } = range;
  if (node.nodeType !== 1 || node.closest('[contenteditable="true"]') !== node) return range;
  let next = node.childNodes[offset] || node.childNodes[offset - 1];
  if (next?.tagName === 'TABLE') next = next.querySelector('td[contenteditable="true"] :is(p, h1, h2, h3, h4, h5, h6)') || next;
  if (!next || next.nodeType !== 1 || next.tagName === 'TABLE') return range;
  const r = document.createRange();
  r.setStart(next, 0);
  r.collapse(true);
  return r;
}

/** Selects a change (an inline one, or the end of a paragraph whose break it is) and shows it. */
function selectChange(item) {
  const range = itemRange(item);
  page.value.focus();
  const sel = window.getSelection();
  sel.removeAllRanges();
  sel.addRange(range);
  savedRange = range.cloneRange();
  item.el.scrollIntoView?.({ block: 'nearest' });
  flash(item.el);
  const rev = revisionOfItem(item);
  message.value = `${describeChange(item).replace(/^./, (c) => c.toUpperCase())}${rev?.date ? `, ${rev.date.slice(0, 10)}` : ''}`;
}

function gotoChange(dir) {
  const items = changeItems(page.value);
  const item = adjacentChange(selectionInPage() || savedRange, items, dir);
  if (item) selectChange(item);
  else message.value = 'There are no tracked changes to go to';
}

/**
 * Accepts or rejects the changes at the selection (see changesAt), as one undoable step. `then`:
 * 'next' goes on to the next change, as Word's Accept and Reject buttons do; 'stay' doesn't.
 */
function resolveCurrent(accept, then = 'next') {
  const range = selectionInPage() || savedRange;
  const items = changeItems(page.value);
  const targets = changesAt(range, items);
  if (!targets.length) {
    message.value = 'Put the cursor on a tracked change (or select some), or use Next change';
    return;
  }
  const what = targets.length === 1 ? describeChange(targets[0]) : `${targets.length} changes`;
  // Afterwards, the caret goes where the first change was.
  const first = targets[0];
  const at = document.createRange();
  if (first.kind === 'mark') at.setStart(first.el, first.el.childNodes.length);
  else at.setStartBefore(['row', 'cells', 'merge', 'tprops', 'rprops', 'cprops'].includes(first.kind) ? first.el.closest('table') : first.el);
  at.collapse(true);
  resolveItems(targets, accept);
  const sel = window.getSelection();
  if (at.startContainer.isConnected) {
    const caret = settleCaret(at);
    sel.removeAllRanges();
    sel.addRange(caret);
    savedRange = caret.cloneRange();
  }
  message.value = `${accept ? 'Accepted' : 'Rejected'} ${what}`;
  if (then === 'next' && targets.length === 1) {
    const next = adjacentChange(selectionInPage(), changeItems(page.value), 1);
    if (next) {
      selectChange(next);
      message.value = `${accept ? 'Accepted' : 'Rejected'} ${what}. Next: ${describeChange(next)}`;
    } else {
      message.value = `${accept ? 'Accepted' : 'Rejected'} ${what}. No more tracked changes`;
    }
  }
}

function resolveAll(accept) {
  const items = changeItems(page.value);
  if (!items.length) return;
  resolveItems(items, accept);
  const word = changes.wordOnly;
  message.value = `${accept ? 'Accepted' : 'Rejected'} all ${items.length} ${items.length === 1 ? 'change' : 'changes'}`
    + (word ? `. ${word} more (in locked content) can only be ${accept ? 'accepted' : 'rejected'} in Word` : '');
}

/** Resolves changes, last first so earlier ones stay where they are, as one undoable step. */
function resolveItems(items, accept) {
  closeChooser();
  undoStack.flush();
  for (const item of [...items].reverse()) {
    if (item.el.isConnected) doc.resolveRevision(item, accept);
  }
  resolved[accept ? 'accepted' : 'rejected'] += items.length;
  doc.renumber(page.value);
  undoStack.commit();
  onSelectionChange();
  markDirty();
  refreshChanges();
  refreshComments();
}

function openChangeMenu(which) {
  if (changeMenu.open === which) return closeChangeMenu();
  const button = which === 'accept' ? acceptMenuButton.value : rejectMenuButton.value;
  const rect = button.getBoundingClientRect();
  Object.assign(changeMenu, { open: which, x: Math.max(8, Math.min(rect.left - 150, window.innerWidth - 220)), y: rect.bottom + 4 });
  nextTick(() => changeMenuEl.value?.querySelector('button')?.focus());
}

function closeChangeMenu() {
  changeMenu.open = null;
  page.value.focus();
  restoreSelection();
}

function menuResolve(accept, how) {
  changeMenu.open = null;
  page.value.focus();
  restoreSelection();
  if (how === 'all') resolveAll(accept);
  else resolveCurrent(accept, how === 'one' ? 'next' : 'stay');
}

// ------------------------------------------------------------------ superscript, subscript, highlight

/** Runs a change to character formatting of the selection as one undoable step; false if refused. */
function formatSelection(change) {
  restoreSelection();
  const range = selectionInPage();
  if (!range || inPlainField()) return false;
  const locked = readOnlySelection();
  if (locked) {
    explainReadOnly(locked, true);
    return false;
  }
  undoStack.flush();
  const after = change(range);
  if (after) {
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(after);
  }
  undoStack.commit();
  onSelectionChange();
  markDirty();
  noteUntrackedFormatting();
  return true;
}

/** Superscript or subscript on or off, as Word's buttons: one turns the other off. */
function toggleVertAlign(kind) {
  formatSelection((range) => {
    if (range.collapsed) {
      document.execCommand(kind); // for what is typed next
      return null;
    }
    const on = formatState(range)[kind];
    return setVertAlign(range, on ? null : kind);
  });
}

// The highlight colour last chosen, remembered in this browser only.
const HIGHLIGHT_KEY = 'intact.highlight';
const highlightColor = ref((() => {
  try {
    const c = localStorage.getItem(HIGHLIGHT_KEY);
    return HIGHLIGHT_COLORS.some(([name]) => name === c) ? c : 'yellow';
  } catch { return 'yellow'; }
})());
const palette = reactive({ open: false, x: 0, y: 0 });
const paletteEl = ref(null);
const paletteButton = ref(null);
const colorLabel = (name) => HIGHLIGHT_COLORS.find(([n]) => n === name)?.[1] || name;

/** Highlights the selection in `color`, or removes highlight (null). `toggle`: off if it is that colour. */
function applyHighlight(color, toggle = false) {
  const range = selectionInPage() || savedRange;
  if (!range || range.collapsed) {
    message.value = 'Select the text to highlight first';
    return;
  }
  formatSelection((r) => {
    const off = !color || (toggle && formatState(r).highlight === color);
    return setHighlight(r, off ? null : color);
  });
}

function togglePalette() {
  if (palette.open) return closePalette();
  const rect = paletteButton.value.getBoundingClientRect();
  Object.assign(palette, { open: true, x: Math.max(8, Math.min(rect.left - 120, window.innerWidth - 200)), y: rect.bottom + 4 });
  nextTick(() => paletteEl.value?.querySelector('.swatch.current, .swatch')?.focus());
}

function closePalette() {
  palette.open = false;
  page.value.focus();
  restoreSelection();
}

function pickHighlight(color) {
  palette.open = false;
  if (color) {
    highlightColor.value = color;
    try { localStorage.setItem(HIGHLIGHT_KEY, color); } catch { /* this session only */ }
  }
  page.value.focus();
  restoreSelection();
  applyHighlight(color);
}

// ------------------------------------------------------------------ keyboard shortcuts

/** Runs one of Word's shortcuts (see editor/shortcuts.js). */
function runShortcut(name) {
  if (['left', 'center', 'right', 'justify'].includes(name)) return alignBlocks(name);
  if (['bold', 'italic', 'underline'].includes(name)) return exec(name);
  if (name === 'normal') return applyBlockStyle('normal');
  if (name.startsWith('heading')) return applyBlockStyle(`h${name.slice(-1)}`);
  if (name === 'bullets') return toggleList('bullet');
  if (name === 'indent' || name === 'outdent') {
    const range = selectionInPage();
    if (range && !selectedBlocks(range).some((b) => doc.listOf(b))) {
      message.value = 'Indenting works on list items here';
      return;
    }
    return changeLevel(name === 'indent' ? 1 : -1);
  }
  if (name === 'clearFormatting') return clearCharacterFormatting();
  if (name === 'superscript' || name === 'subscript') return toggleVertAlign(name);
  if (name === 'highlight') return applyHighlight(highlightColor.value, true);
  if (name === 'formattingMarks') return toggleMarks();
  if (name === 'comment') return startComment();
  if (name === 'nonBreakingSpace') return typeText('\u00a0');
  if (name === 'pageBreak') return insertPageBreak();
}

/**
 * Ctrl+Space: takes bold, italic, underline, strikethrough, superscript and subscript off the
 * selection. Highlight stays, as in Word.
 */
function clearCharacterFormatting() {
  formatSelection((range) => {
    const after = range.collapsed ? null : setVertAlign(range, null);
    if (after) {
      const sel = window.getSelection();
      sel.removeAllRanges();
      sel.addRange(after);
    }
    for (const cmd of CHAR_FORMATS) {
      // A mixed selection may first turn fully on; the second toggle then turns it off.
      for (let i = 0; i < 2 && document.queryCommandState(cmd); i++) document.execCommand(cmd);
    }
    return null;
  });
}

/** Types `text` at the selection, as a key press would: tracked while tracking, never into read-only content. */
function typeText(text) {
  const range = selectionInPage();
  if (!range) return;
  const locked = readOnlySelection();
  if (locked) return explainReadOnly(locked, true);
  if (!range.collapsed && refuseRemoving([range])) return;
  undoStack.flush();
  const opts = trackedBy(range.startContainer);
  if (opts) {
    if (!range.collapsed) deleteTracked(range, false, opts);
    if (typeTracked(text, opts)) onInput();
    else document.execCommand('insertText', false, text);
  } else {
    if (!range.collapsed) deleteSelection();
    if (typeUntracked(text)) onInput();
    else document.execCommand('insertText', false, text);
  }
}

/** Ctrl+Enter: a page break at the caret (tracked while tracking). Not inside form fields. */
function insertPageBreak() {
  const range = selectionInPage();
  if (!range) return;
  if (inField(range.startContainer)) {
    message.value = 'A page break can’t go inside a field';
    return;
  }
  const locked = readOnlySelection();
  if (locked) return explainReadOnly(locked, true);
  if (!range.collapsed && refuseRemoving([range])) return;
  undoStack.flush();
  const opts = trackedBy(range.startContainer);
  if (!range.collapsed) {
    if (opts) deleteTracked(range, false, opts);
    else deleteSelection();
  }
  const pageBreak = doc.newPageBreak();
  if (opts) insertTracked([pageBreak], opts);
  else insertAtCaret([pageBreak]);
  onInput();
}

// ------------------------------------------------------------------ track changes

let pendingSplit = null; // the author of a paragraph break being inserted while tracking

/** The author and date for a tracked edit made now, or null when this edit isn't tracked. */
function trackedBy(node) {
  return tracking.value && node && !inField(node) ? { author: author.value, date: revisionDate() } : null;
}

/**
 * Typing, deleting, Enter and Shift+Enter while tracking: recorded as tracked changes instead of
 * changing the text (see editor/track.js). True if the event was dealt with here.
 */
function trackedBeforeInput(e) {
  const range = selectionInPage();
  const opts = trackedBy(range?.startContainer);
  if (!opts) {
    // Track Changes off: text typed in (or right after) one of this session's insertions is plain.
    if (e.inputType === 'insertText' && e.data && range?.collapsed && typeUntracked(e.data)) {
      e.preventDefault();
      onInput();
      return true;
    }
    return false;
  }
  const type = e.inputType;
  if (type === 'insertFromDrop' || type === 'deleteByDrag') {
    e.preventDefault();
    message.value = 'Dragging text isn’t tracked – cut and paste it instead while tracking changes';
    return true;
  }
  if (type === 'insertCompositionText') return true; // wrapped as an insertion on compositionend
  const deleting = type.startsWith('delete');
  if (!deleting && !['insertText', 'insertReplacementText', 'insertParagraph', 'insertLineBreak'].includes(type)) return false;
  let removal = null;
  if (!range.collapsed) removal = range.cloneRange();
  else if (type === 'deleteContentBackward' || type === 'deleteContentForward') removal = unitRange(type === 'deleteContentBackward');
  else if (deleting) removal = rangesRemovedBy(e)[0] || null;
  if (deleting && !removal) {
    e.preventDefault(); // at the very start (or end): nothing to delete
    return true;
  }
  if (removal && refuseRemoving([removal])) {
    e.preventDefault();
    return true;
  }
  if (type === 'insertParagraph' && !removal) {
    pendingSplit = opts; // the browser splits the paragraph; onInput marks the new break
    return true;
  }
  if (type === 'insertText' && !removal && !typeTracked(e.data || '', opts)) return true; // extends an insertion
  e.preventDefault();
  if (removal) deleteTracked(removal, /Backward$/.test(type), opts);
  if (type === 'insertText' || type === 'insertReplacementText') {
    const text = e.data ?? e.dataTransfer?.getData('text/plain') ?? '';
    if (removal && text && !typeTracked(text, opts)) {
      document.execCommand('insertText', false, text); // into an insertion; fires input
      return true;
    }
  } else if (type === 'insertParagraph') {
    pendingSplit = opts;
    document.execCommand('insertParagraph'); // fires input, which marks the break
    return true;
  } else if (type === 'insertLineBreak') {
    document.execCommand('insertLineBreak');
    wrapInsertedBreak(opts);
    undoStack.schedule();
    return true;
  }
  onInput();
  return true;
}

/** Text from an input method can't be intercepted: it becomes an insertion once it is complete. */
function onCompositionEnd(e) {
  const range = selectionInPage();
  const opts = trackedBy(range?.startContainer);
  if (!opts) return;
  wrapComposedText(e.data, opts);
  undoStack.schedule();
}

function noteUntrackedFormatting() {
  if (tracking.value) message.value = 'Formatting isn’t tracked – it’s applied directly';
}

function toggleTracking() {
  if (tracking.value) {
    tracking.value = false;
    message.value = 'Track Changes is off: your edits change the text directly';
  } else if (named.value) {
    startTracking();
  } else {
    openAuthorDialog(true);
  }
}

function startTracking() {
  tracking.value = true;
  message.value = `Track Changes is on: your edits are marked as ${author.value}’s`;
  page.value.focus();
  restoreSelection();
}

async function openAuthorDialog(start = false) {
  start = start === true; // not a click event
  const anchor = (start ? trackButton.value : document.querySelector('.tracking-status')) || trackButton.value;
  const rect = anchor.getBoundingClientRect();
  Object.assign(authorDialog, {
    open: true,
    start,
    name: named.value ? author.value : '',
    x: Math.max(8, Math.min(rect.left, window.innerWidth - 300)),
    y: start ? rect.bottom + 4 : Math.max(8, rect.top - 140),
  });
  await nextTick();
  authorInput.value?.focus();
  authorInput.value?.select();
}

function closeAuthorDialog() {
  authorDialog.open = false;
  page.value.focus();
  restoreSelection();
}

function confirmAuthor() {
  const name = authorDialog.name.trim();
  if (!name) return;
  author.value = name;
  try { localStorage.setItem(AUTHOR_KEY, name); } catch { /* remembered for this session only */ }
  named.value = true;
  authorDialog.open = false;
  if (authorDialog.start) startTracking();
  else {
    message.value = `Your new tracked changes are marked as ${name}’s`;
    page.value.focus();
    restoreSelection();
  }
}

/** Deletes the selection with execCommand (which fires no beforeinput), keeping comment ranges. */
function deleteSelection() {
  const before = commentMarkers(page.value);
  document.execCommand('delete');
  restoreCommentMarkers(before, page.value);
}

/**
 * Refuses an edit that would remove a tracked change, saying why: true if `ranges` (what the edit
 * would delete) touch one. Edits made with execCommand fire no beforeinput, so call this first.
 */
function refuseRemoving(ranges) {
  const el = trackedChangeAt(ranges, page.value);
  if (!el) return false;
  message.value = refusalMessage(el);
  flash(el);
  return true;
}

/** Flashes what an edit can't change. An animation, not a class, so undo snapshots don't change. */
function flash(el) {
  el.animate?.([
    { outline: '2px solid #b3261e', outlineOffset: '1px', backgroundColor: 'rgb(179 38 30 / 15%)' },
    { outline: '2px solid transparent', outlineOffset: '1px' },
  ], { duration: 1200, easing: 'ease-out' });
}

// ------------------------------------------------------------------ read-only content

// The read-only block or cell last clicked. The editor gives up focus then (the browser would make
// up a caret for the next key otherwise), and keys typed next say why nothing happens.
let lastReadOnly = null;
let pointerDown = null;

/** The read-only element the selection starts or ends in, or null. */
function readOnlySelection() {
  const range = selectionInPage();
  return range && (readOnlyAt(range.startContainer, page.value) || readOnlyAt(range.endContainer, page.value));
}

/** Says in the status bar what read-only `el` is and where to edit it; `refused` flashes it too. */
function explainReadOnly(el, refused = false) {
  message.value = readOnlyMessage(el, page.value);
  if (refused) flash(el);
}

function onPageMousedown(e) {
  const pic = e.target.closest?.('.picture');
  if (pic && page.value.contains(pic)) {
    onPictureMousedown(e, pic);
    return;
  }
  pointerDown = { x: e.clientX, y: e.clientY, range: selectionInPage()?.cloneRange() || null };
}

/**
 * A click (not a drag) on read-only content says what it is. The browser leaves the caret where it
 * was, so typing would land there, out of sight: put it beside inline content, on the side clicked,
 * or clear it for a block or cell.
 */
function clickReadOnly(el, e) {
  const down = pointerDown;
  pointerDown = null;
  if (down && Math.hypot(e.clientX - down.x, e.clientY - down.y) > 4) return;
  explainReadOnly(el);
  const now = selectionInPage();
  const moved = now && down?.range && !(now.startContainer === down.range.startContainer && now.startOffset === down.range.startOffset
    && now.endContainer === down.range.endContainer && now.endOffset === down.range.endOffset);
  if (moved && !readOnlyAt(now.startContainer, page.value)) return; // the browser placed the caret itself
  const sel = window.getSelection();
  const beside = caretBeside(el, e.clientX, e.clientY);
  sel.removeAllRanges();
  lastReadOnly = beside ? null : el;
  if (!beside) {
    document.activeElement?.blur();
    return;
  }
  el.parentElement.closest('[contenteditable="true"]')?.focus({ preventScroll: true });
  sel.addRange(beside);
}

function onHeaderFooterClick(e) {
  if (!e.target.closest('button')) message.value = `Read-only: ${e.currentTarget.title}`;
}

/**
 * Typing while the selection is in a header, footer or comment (outside the editable page), or
 * after clicking a read-only block, does nothing; say why. True if the key was explained.
 */
function explainKeyOutsideText(e) {
  if (!isEditingKey(e) || document.activeElement?.matches('input, textarea, select, [contenteditable="true"]')) return false;
  if (lastReadOnly?.isConnected) {
    explainReadOnly(lastReadOnly, true);
    return true;
  }
  const node = window.getSelection().anchorNode;
  const el = node?.nodeType === 1 ? node : node?.parentElement;
  const hf = el?.closest('.docx-hf');
  if (hf) {
    message.value = `Read-only: ${hf.title}`;
    flash(hf);
    return true;
  }
  if (el?.closest('.comments') && e.key !== 'Enter') {
    message.value = 'Read-only: comments are shown for reference – reply to or resolve them in Word';
    return true;
  }
  return false;
}

function onKeydown(e) {
  if (e.key === 'Escape' && picture.cropping) {
    e.preventDefault();
    toggleCropping();
    return;
  }
  const mod = e.ctrlKey || e.metaKey;
  const key = e.key.toLowerCase();
  const range = selectionInPage();
  if (isEditingKey(e) || (e.key === 'Tab' && !mod && !e.altKey)) {
    const locked = readOnlySelection();
    if (locked) {
      e.preventDefault();
      explainReadOnly(locked, true);
      return;
    }
  }
  const field = range && controlAt(range.startContainer, page.value)?.querySelector(':scope > .sdt-content');
  if ((e.altKey && !mod && e.key === 'ArrowDown') || (e.key === ' ' && !mod && !e.altKey && range && !range.collapsed)) {
    const near = fieldAtSelection(range, page.value);
    if (near) {
      e.preventDefault();
      activateField(near);
      return;
    }
  }
  if (mod && !e.altKey && ['b', 'i', 'u'].includes(key) && inPlainField()) {
    e.preventDefault();
    return;
  }
  const shortcut = shortcutFor(e);
  if (shortcut) {
    e.preventDefault();
    runShortcut(shortcut);
    return;
  }
  if (e.key === 'Enter' && field?.tagName === 'SPAN') {
    // A field inside a line can't hold paragraphs; multi-line plain text fields take line breaks.
    e.preventDefault();
    if (e.shiftKey || field.dataset.plain === 'multi') document.execCommand('insertLineBreak');
    return;
  }
  if (mod && key === 's') {
    e.preventDefault();
    requestDownload();
  } else if (mod && !e.altKey && (key === 'z' || key === 'y')) {
    e.preventDefault();
    if (key === 'y' || e.shiftKey) redo();
    else undo();
  } else if (e.key === 'Tab' && !mod && !e.altKey) {
    e.preventDefault();
    onTab(e);
  } else if ((e.key === 'Enter' || e.key === 'Backspace') && !mod && !e.altKey && !e.shiftKey && !e.isComposing) {
    // Word: Enter on an empty list item, or Backspace at the start of one, ends the item's list.
    const range = selectionInPage();
    const block = range?.collapsed && closestBlock(range.startContainer);
    if (!block || !doc.listOf(block)) return;
    if (e.key === 'Enter' ? !isEmptyBlock(block) : !isAtBlockStart(range, block)) return;
    e.preventDefault();
    if (e.key === 'Enter') endListItem(block);
    else changeLists(() => doc.removeList(block));
  }
}

function onPaste(e) {
  e.preventDefault();
  const range = selectionInPage();
  const locked = readOnlySelection();
  if (locked) {
    explainReadOnly(locked, true);
    return;
  }
  if (!range || refuseRemoving([range])) return;
  // A picture on its own (a screenshot, an image copied from an app): inserted as a picture.
  // Some browsers list a copied picture only among the clipboard's items, not its files.
  const files = e.clipboardData.files?.length ? [...e.clipboardData.files]
    : [...(e.clipboardData.items || [])].filter((i) => i.kind === 'file').map((i) => i.getAsFile()).filter(Boolean);
  const pictures = files.filter((f) => f.type.startsWith('image/'));
  if (pictures.length && !e.clipboardData.getData('text/plain').trim()) {
    insertPictures(pictures);
    return;
  }
  undoStack.flush();
  const html = e.clipboardData.getData('text/html');
  let text = e.clipboardData.getData('text/plain').replace(/\r\n?/g, '\n');
  const field = controlAt(range.startContainer, page.value)?.querySelector(':scope > .sdt-content');
  if (field?.tagName === 'SPAN' || field?.dataset.plain) {
    // Fields inside a line, and plain-text fields, take text only; inside a line it stays on one
    // line (or becomes line breaks in multi-line fields).
    if (field.tagName === 'SPAN' && field.dataset.plain !== 'multi') text = text.replace(/\s*\n\s*/g, ' ');
    document.execCommand('insertText', false, text);
    undoStack.commit();
    onInput();
    return;
  }
  const { blocks, skippedImages } = parseClipboard(html, text);
  if (!blocks.length) return;
  const tracked = trackedBy(range.startContainer);
  if (!range.collapsed) {
    if (tracked) deleteTracked(range, false, tracked);
    else deleteSelection();
  }
  pasteBlocks(blocks, tracked);
  onInput();
  undoStack.commit();
  if (skippedImages) {
    message.value = `${skippedImages} ${skippedImages === 1 ? 'picture was' : 'pictures were'} not pasted with the text – copy a picture on its own and paste it, or use Insert picture`;
  }
}

/**
 * Inserts parsed blocks with Word's "merge formatting": text takes the formatting where it lands plus
 * its own bold/italic/underline/strike. The first block joins the paragraph at the caret; each
 * further one starts a new paragraph that keeps the destination paragraph's settings, unless it is
 * a list item, which joins the destination list if it is of the same kind or else a new list.
 */
function pasteBlocks(blocks, tracked = null) {
  const startRange = selectionInPage();
  const dest = startRange && closestBlock(startRange.startContainer);
  const destWasEmpty = !!dest && isEmptyBlock(dest);
  const destNum = dest ? { numId: dest.dataset.numId, ilvl: dest.dataset.ilvl } : {};
  const destList = dest ? doc.listOf(dest) : null;
  const destKind = doc.listKind(destList);
  const newLists = {};

  const applyList = (el, block) => {
    if (block.list) {
      const { kind, level } = block.list;
      const numId = destList && destKind === kind ? destList.numId : (newLists[kind] ||= doc.newList(kind));
      doc.setList(el, numId, doc.hasListLevel(numId, level) ? level : 0);
    } else if (destNum.numId != null) {
      el.dataset.numId = destNum.numId;
      el.dataset.ilvl = destNum.ilvl ?? '0';
    } else {
      delete el.dataset.numId;
      delete el.dataset.ilvl;
    }
  };

  blocks.forEach((block, i) => {
    if (i > 0) {
      document.execCommand('insertParagraph');
      if (tracked) markSplit(tracked);
    }
    const here = selectionInPage();
    const el = here && closestBlock(here.startContainer);
    if (el && (i > 0 || destWasEmpty)) applyList(el, block);
    if (tracked && block.runs.length) insertTracked(runNodes(block.runs), tracked);
    else insertAtCaret(runNodes(block.runs));
  });
  doc.renumber(page.value);
}

/** Picture files dropped on the page are inserted where they are dropped; other files aren't. */
function onDrop(e) {
  const files = [...(e.dataTransfer?.files || [])];
  if (!files.length) return;
  e.preventDefault();
  const pictures = files.filter((f) => f.type.startsWith('image/'));
  if (!pictures.length) return;
  let at = null;
  if (document.caretPositionFromPoint) {
    const pos = document.caretPositionFromPoint(e.clientX, e.clientY);
    if (pos) {
      at = document.createRange();
      at.setStart(pos.offsetNode, pos.offset);
    }
  } else if (document.caretRangeFromPoint) {
    at = document.caretRangeFromPoint(e.clientX, e.clientY);
  }
  insertPictures(pictures, at && page.value.contains(at.startContainer) ? at : null);
}

function onClick(e) {
  if (e.target.closest('.picture')) return; // selected on mousedown
  const marker = e.target.closest('.comment-ref');
  if (marker) {
    commentsOpen.value = true;
    selectComment(marker.dataset.commentId, false);
    return;
  }
  const placeholderCell = e.target.closest('td.sdt-cell.showing-placeholder');
  if (placeholderCell) {
    selectCell(placeholderCell); // a field around a table cell: typing replaces its placeholder
    return;
  }
  const locked = readOnlyAt(e.target, page.value);
  if (locked) {
    clickReadOnly(locked, e);
    return;
  }
  const control = controlAt(e.target, page.value);
  if (control?.classList.contains('showing-placeholder') && e.target.closest('.sdt-content')) {
    selectCell(controlBox(control)); // the browser resets the selection on mouseup
    return;
  }
  if (control && !e.target.closest('.sdt-content')) {
    const kind = doc.control(control).kind;
    if (kind === 'checkbox') toggleCheckbox(control);
    else if (kind !== 'text') openChooser(control);
    return;
  }
  const a = e.target.closest('a[href]');
  if (a && (e.ctrlKey || e.metaKey) && /^(https?:|mailto:)/i.test(a.getAttribute('href'))) {
    e.preventDefault();
    window.open(a.href, '_blank', 'noopener');
  }
}

// ------------------------------------------------------------------ content controls

/** Applies a change to a control as one undoable step. */
function changeControl(el, change) {
  undoStack.flush();
  change();
  clearPlaceholder(el, (id) => doc.isPlaceholderStyle(id));
  undoStack.commit();
  markDirty();
}

function toggleCheckbox(el) {
  const { symbols } = doc.control(el);
  changeControl(el, () => {
    el.dataset.checked = el.dataset.checked === '1' ? '0' : '1';
    setControlText(el, el.dataset.checked === '1' ? symbols.on : symbols.off);
  });
}

function openChooser(el) {
  const info = doc.control(el);
  const box = el.getBoundingClientRect();
  chooserTarget = el;
  Object.assign(chooser, {
    open: true,
    kind: info.kind,
    x: Math.max(8, Math.min(box.left, window.innerWidth - 240)),
    y: box.bottom + 4,
    choices: info.choices,
    current: el.classList.contains('showing-placeholder') ? '' : controlBox(el).textContent,
    date: info.date?.iso || new Date().toISOString().slice(0, 10),
  });
  nextTick(() => (dateInput.value || chooserEl.value?.querySelector('[aria-selected="true"], button'))?.focus());
}

function closeChooser(refocus = false) {
  const el = chooserTarget;
  chooser.open = false;
  chooserTarget = null;
  if (refocus) {
    // Back to the document, with the caret just after the field.
    page.value.focus();
    if (el?.isConnected) {
      const range = document.createRange();
      range.setStartAfter(el);
      const sel = window.getSelection();
      sel.removeAllRanges();
      sel.addRange(range);
    }
  }
}

/** ↑ / ↓ / Home / End move between the items of a drop-down list. */
function onChooserKeydown(e) {
  if (chooser.kind === 'date' || !['ArrowDown', 'ArrowUp', 'Home', 'End'].includes(e.key)) return;
  e.preventDefault();
  const items = [...chooserEl.value.querySelectorAll('button')];
  const i = items.indexOf(document.activeElement);
  const next = { ArrowDown: Math.min(i + 1, items.length - 1), ArrowUp: Math.max(i - 1, 0), Home: 0, End: items.length - 1 }[e.key];
  items[next]?.focus();
}

// A date typed digit by digit is "complete" (and fires change) before the year is finished, so typed
// dates are taken on Enter; a date picked from the calendar is taken at once.
let lastDateKey = 0;
function onDateKeydown(e) {
  lastDateKey = Date.now();
  if (e.key === 'Enter') {
    e.preventDefault();
    chooseDate(e.target.value);
  }
}
function onDateChange(e) {
  if (Date.now() - lastDateKey > 400) chooseDate(e.target.value);
}

/** Alt+↓, or Space on a selected field: tick a check box, or open a list or date picker. */
function activateField(el) {
  if (doc.control(el).kind === 'checkbox') toggleCheckbox(el);
  else openChooser(el);
}

const canMarkFields = typeof CSS !== 'undefined' && CSS.highlights && typeof Highlight !== 'undefined';

/** Marks the field the caret is touching, so it's clear that Alt+↓ will act on it. */
function markNearField(range) {
  if (!canMarkFields) return;
  const field = fieldAtSelection(range, page.value);
  if (!field) {
    CSS.highlights.delete('field-near');
    return;
  }
  const r = document.createRange();
  r.selectNodeContents(field);
  CSS.highlights.set('field-near', new Highlight(r));
}

function onDocumentMousedown(e) {
  if (chooser.open && !chooserEl.value?.contains(e.target) && !chooserTarget?.contains(e.target)) closeChooser();
  if (palette.open && !paletteEl.value?.contains(e.target) && !paletteButton.value?.contains(e.target)) palette.open = false;
  if (picturePanel.open && !picturePanelEl.value?.contains(e.target) && !layoutOptionsButton.value?.contains(e.target)) picturePanel.open = false;
  if (splitDialog.open && !splitDialogEl.value?.contains(e.target)) splitDialog.open = false;
  if (tableMenu.open && !tableMenuEl.value?.contains(e.target) && !layoutButton.value?.contains(e.target)) tableMenu.open = false;
  if (tablePicker.open && !tablePickerEl.value?.contains(e.target) && !tableButton.value?.contains(e.target)) tablePicker.open = false;
  if (changeMenu.open && !changeMenuEl.value?.contains(e.target) && !acceptMenuButton.value?.contains(e.target) && !rejectMenuButton.value?.contains(e.target)) changeMenu.open = null;
}

function choose(choice) {
  const el = chooserTarget;
  closeChooser(true);
  if (!el?.isConnected) return;
  changeControl(el, () => {
    setControlText(el, choice.text);
    el.dataset.value = choice.value;
  });
}

function chooseDate(iso) {
  const el = chooserTarget;
  if (!iso || !el?.isConnected) return;
  closeChooser(true);
  changeControl(el, () => {
    setControlText(el, doc.formatDate(el, iso));
    el.dataset.fullDate = iso;
  });
}

// ------------------------------------------------------------------ comments

/**
 * Top-level comments in the order they appear in the text, each with its replies, leaving out those
 * whose text was deleted (ids in `removed`).
 */
function commentThreads(removed = new Set()) {
  // As shown: with the text and resolved state kept on their marks in the page (see commentMark).
  const comments = doc.comments.filter((c) => !removed.has(c.id)).map((c) => {
    const mark = commentMark(c.id);
    const text = mark?.dataset.edit;
    return {
      ...c,
      anchored: !!mark,
      paragraphs: text != null ? text.split('\n').map((s) => s.trim()).filter(Boolean) : c.paragraphs,
      done: mark?.dataset.done != null ? mark.dataset.done === '1' : c.done,
    };
  });
  const ids = new Set(comments.map((c) => c.id));
  const order = (c) => doc.commentOrder.get(c.id) ?? Infinity;
  return comments
    .filter((c) => !c.parentId || !ids.has(c.parentId))
    .sort((a, b) => order(a) - order(b))
    .map((c) => ({ ...c, n: doc.commentOrder.get(c.id) || null, replies: comments.filter((r) => r.parentId === c.id) }));
}

// As in Word, deleting the text a comment is attached to deletes the comment. The panel follows the
// text, and the status bar says so, since Undo is the way back.
let removedComments = new Set();

/** A comment's mark in the page: it carries the comment's edited text and resolved state. */
function commentMark(id) {
  return page.value?.querySelector(`[data-comment-ref="${CSS.escape(id)}"]`) || null;
}

function refreshComments() {
  if (!doc.comments.length) return;
  doc.renumberComments(page.value); // marks may have been added, removed or brought back by Undo
  const removed = doc.removedComments(page.value);
  const added = [...removed].filter((id) => !removedComments.has(id));
  removedComments = removed;
  threads.value = commentThreads(removed); // edits and resolved state live on the marks, so always
  paintComments();
  if (added.length) message.value = commentRemovalMessage(added);
}

function commentRemovalMessage(ids) {
  const removed = doc.comments.filter((c) => ids.includes(c.id));
  const top = removed.filter((c) => !c.parentId || !ids.includes(c.parentId));
  const what = top.length === 1
    ? `the comment by ${top[0].author}${removed.length > 1 ? ` and ${removed.length - 1 === 1 ? 'its reply' : `its ${removed.length - 1} replies`}` : ''}`
    : `${removed.length} comments`;
  return `Deleted ${what} with the text ${removed.length === 1 ? 'it was' : 'they were'} attached to – Undo (Ctrl+Z) brings ${removed.length === 1 ? 'it' : 'them'} back`;
}

/** The commented text: from the range start marker to the end marker (or the reference mark). */
function commentRange(id) {
  const q = (sel) => page.value.querySelector(`[${sel}="${CSS.escape(id)}"]`);
  const start = q('data-comment-start');
  const end = q('data-comment-end');
  const ref = q('data-comment-id');
  if (!start && !ref) return null;
  const range = document.createRange();
  if (start) range.setStartAfter(start);
  else range.setStartBefore(ref);
  if (end) range.setEndBefore(end);
  else range.setEndAfter(ref);
  return range;
}

const canHighlight = typeof CSS !== 'undefined' && CSS.highlights && typeof Highlight !== 'undefined';

function paintComments() {
  if (!canHighlight) return;
  const all = new Highlight();
  const active = new Highlight();
  active.priority = 1;
  for (const t of threads.value) {
    const range = commentRange(t.id);
    if (!range) continue;
    all.add(range);
    if (t.id === activeComment.value) active.add(range);
  }
  if (composer.open && composerRange) active.add(composerRange); // the text a new comment will be on
  CSS.highlights.set('comment', all);
  CSS.highlights.set('comment-active', active);
}

function selectComment(id, reveal = true) {
  activeComment.value = id;
  paintComments();
  if (reveal) {
    const range = commentRange(id);
    if (range) revealRect(range.getBoundingClientRect());
  } else {
    nextTick(() => document.querySelector('.comments .thread.active')?.scrollIntoView({ block: 'nearest' }));
  }
}

function revealRect(rect) {
  const canvas = page.value.closest('.canvas');
  const box = canvas.getBoundingClientRect();
  if (rect.top < box.top + 8 || rect.bottom > box.bottom - 8) canvas.scrollTop += rect.top - box.top - box.height / 3;
}

// ------------------------------------------------------------------ find and replace

/** Ctrl+F / Ctrl+H anywhere in the editor open Find (and Replace) instead of the browser's own. */
function onWindowKeydown(e) {
  if (explainKeyOutsideText(e)) return;
  const mod = e.ctrlKey || e.metaKey;
  const key = e.key.toLowerCase();
  if (mod && e.shiftKey && !e.altKey && key === 'e') {
    e.preventDefault(); // Word's Track Changes shortcut, wherever the focus is in the editor
    toggleTracking();
    return;
  }
  if (!mod || e.altKey || e.shiftKey || (key !== 'f' && key !== 'h')) return;
  e.preventDefault();
  openFind(key === 'h');
}

async function openFind(replace) {
  const range = selectionInPage() || savedRange;
  const selected = range && !range.collapsed ? range.toString() : '';
  const text = selected && selected.length <= 200 && !/\n/.test(selected) ? selected : '';
  // Re-pressing Ctrl+F / Ctrl+H while Find is open keeps the current match.
  const from = range && !findOpen.value ? { node: range.startContainer, offset: range.startOffset } : null;
  findOpen.value = true;
  await nextTick();
  findBar.value?.open({ replace, text, from });
}

function closeFind() {
  findOpen.value = false;
  page.value.focus();
}

/** What Find searches, in order: headers, the document, footers, then comment texts. */
function searchRoots() {
  return [
    ...(headerEl.value?.children || []),
    page.value,
    ...(footerEl.value?.children || []),
    ...document.querySelectorAll('.comments .comment .text'),
  ];
}

/** Brings a match into view, switching to its header/footer tab or opening Comments if need be. */
async function revealRange(range) {
  const el = range.startContainer.parentElement;
  const variant = el?.closest('.hf-variant');
  if (variant?.hidden) hf[variant.dataset.which] = Number(variant.dataset.variant);
  if (el?.closest('.comments') && !commentsOpen.value) commentsOpen.value = true;
  await nextTick();
  const scroller = el?.closest('.canvas, .comments');
  if (!scroller) return;
  const rect = range.getBoundingClientRect();
  const box = scroller.getBoundingClientRect();
  if (rect.top < box.top + 8 || rect.bottom > box.bottom - 8) scroller.scrollTop += rect.top - box.top - box.height / 3;
}

/** Closing Find leaves the current match selected in the document, as in Word. */
function onFindClosed(range) {
  closeFind();
  if (range && page.value.contains(range.startContainer)) {
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
  }
}

/** Runs a change made outside the browser's editing (such as Replace) as one undoable edit. */
function applyEdit(change) {
  undoStack.flush();
  const touched = change() || [];
  for (const node of touched) {
    const control = controlAt(node, page.value);
    if (control) clearPlaceholder(control, (id) => doc.isPlaceholderStyle(id));
  }
  page.value.normalize();
  undoStack.commit();
  markDirty();
}

// ------------------------------------------------------------------ file actions

// "Review my changes" before Download: what will go back to the recipient. It can be skipped; the
// choice is remembered in this browser only, and the toolbar's Review button still shows it.
const SKIP_REVIEW_KEY = 'intact.skipReview';
const reviewData = ref(null);
const skipReview = ref((() => { try { return localStorage.getItem(SKIP_REVIEW_KEY) === '1'; } catch { return false; } })());

function requestDownload() {
  if (skipReview.value) save();
  else openReview();
}

function openReview() {
  closeChooser();
  undoStack.flush();
  reviewData.value = doc.review(page.value);
}

function closeReview() {
  reviewData.value = null;
  page.value.focus();
  restoreSelection();
}

function downloadReviewed() {
  reviewData.value = null;
  page.value.focus(); // carry on editing where you were
  restoreSelection();
  save();
}

function setSkipReview(on) {
  skipReview.value = on;
  try { localStorage.setItem(SKIP_REVIEW_KEY, on ? '1' : '0'); } catch { /* this session only */ }
}

async function save() {
  if (saving.value) return;
  saving.value = true;
  try {
    const name = `${baseName.value.trim() || 'document'}.docx`;
    const blob = await doc.save(page.value);
    doc.name = name;
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = name;
    document.body.append(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 10000);
    dirty.value = false;
    const n = doc.commentsRemoved.size;
    message.value = `Saved ${name}${n ? ` – ${n} ${n === 1 ? 'comment' : 'comments'} removed with ${n === 1 ? 'its' : 'their'} text` : ''}`;
  } catch (e) {
    console.error(e);
    message.value = `Save failed: ${e.message}`;
  } finally {
    saving.value = false;
  }
}

async function close() {
  const ok = await router.push({ name: 'home' });
  if (ok === undefined) setDocument(null); // navigation succeeded
}
</script>

<style scoped>
.editor-shell { height: 100vh; height: 100dvh; display: flex; flex-direction: column; }
.file { display: flex; align-items: center; gap: 2px; min-width: 0; flex: 1; }
.filename {
  min-width: 0; max-width: 100%; font: inherit; color: #fff; background: transparent;
  border: 1px solid transparent; border-radius: 4px; padding: 4px 6px;
}
.filename:hover, .filename:focus { border-color: rgb(255 255 255 / 50%); outline: none; background: rgb(255 255 255 / 10%); }
.ext { opacity: .7; }
.dirty { color: #ffd166; margin-left: 6px; }
.primary { background: #fff; color: var(--brand); border: 0; border-radius: 4px; padding: 6px 14px; font-weight: 600; cursor: pointer; }
.primary:disabled { opacity: .6; }
@media (max-width: 560px) { .ext, :deep(.nav) { display: none; } }
.ghost { background: transparent; color: #fff; border: 1px solid rgb(255 255 255 / 50%); border-radius: 4px; padding: 5px 12px; cursor: pointer; }

.toolbar {
  display: flex; flex-wrap: wrap; align-items: center; gap: 2px; padding: 6px 12px;
  background: var(--surface); border-bottom: 1px solid var(--border);
}
.toolbar button {
  min-width: 30px; height: 32px; display: inline-flex; align-items: center; justify-content: center;
  border: 1px solid transparent; border-radius: 4px; background: transparent; color: inherit; font-size: 15px; cursor: pointer;
}
.toolbar button:hover:not(:disabled) { background: var(--hover); }
.toolbar button:disabled { opacity: .4; cursor: default; }
.toolbar button.on { background: var(--active); border-color: var(--border); }
.toolbar select { height: 32px; min-width: 130px; border: 1px solid var(--border); border-radius: 4px; background: var(--surface); color: inherit; padding: 0 6px; }
.sep { width: 1px; height: 22px; background: var(--border); margin: 0 4px; }

.workspace { flex: 1; min-height: 0; display: flex; }
.canvas { flex: 1; min-width: 0; overflow: auto; background: var(--canvas); padding: 24px 16px 48px; }
.docx-sheet {
  position: relative; isolation: isolate; /* pictures behind the text stay on the page */ box-sizing: border-box; margin: 0 auto; max-width: 100%; background: #fff; color: #000;
  box-shadow: 0 1px 3px rgb(0 0 0 / 20%), 0 4px 16px rgb(0 0 0 / 8%);
  display: flex; flex-direction: column;
}
.docx-page {
  flex: 1; box-sizing: border-box; outline: none;
  white-space: pre-wrap; overflow-wrap: break-word; tab-size: 4; caret-color: #000;
}
.docx-hf { flex: none; white-space: pre-wrap; overflow-wrap: break-word; tab-size: 4; opacity: .55; cursor: default; }
.docx-hf.footer { display: flex; flex-direction: column; justify-content: flex-end; }
.docx-hf:hover { opacity: .8; }
.docx-hf.empty { padding-bottom: 0 !important; padding-top: 0 !important; }
.docx-hf { position: relative; }
.hf-tabs {
  position: absolute; left: 8px; right: 8px; display: flex; flex-wrap: wrap; gap: 4px;
  font: 11px/1.4 system-ui, sans-serif; white-space: normal;
}
.header .hf-tabs { top: 6px; }
.footer .hf-tabs { bottom: 6px; }
.hf-tabs button {
  padding: 1px 8px; border: 1px solid #c9d1de; border-radius: 9px; background: #fff; color: #44516a; cursor: pointer;
}
.hf-tabs button[aria-selected="true"] { background: #2b579a; border-color: #2b579a; color: #fff; }
.comments-toggle { gap: 3px; padding: 0 6px; }
.script { font-size: 14px; line-height: 1; }
.script sup, .script sub { font-size: 9px; }
.split { display: inline-flex; }
.split .arrow { min-width: 16px !important; padding: 0 2px; font-size: 10px !important; }
.changes-group { display: inline-flex; align-items: center; gap: 2px; }
.labelled { gap: 4px; padding: 0 6px !important; font-size: 13px !important; }
@media (max-width: 1600px) { .labelled .label, .has-changes .review-toggle .label, .has-changes .track-toggle .label { display: none; } }
.change-menu {
  position: fixed; z-index: 10; min-width: 200px; padding: 4px; display: flex; flex-direction: column;
  background: var(--surface); border: 1px solid var(--border); border-radius: 6px; box-shadow: 0 4px 16px rgb(0 0 0 / 15%);
}
.change-menu button { text-align: left; padding: 6px 10px; border: 0; border-radius: 4px; background: transparent; color: inherit; font: inherit; cursor: pointer; }
.change-menu button:hover, .change-menu button:focus-visible { background: var(--hover); }
.picture-frame { position: absolute; z-index: 6; border: 1px solid var(--accent); pointer-events: none; box-sizing: border-box; }
.picture-frame .handle { position: absolute; width: 9px; height: 9px; margin: -5px 0 0 -5px; background: #fff; border: 1px solid var(--accent); border-radius: 50%; pointer-events: auto; touch-action: none; }
.picture-frame .nw { left: 0; top: 0; cursor: nwse-resize; }
.picture-frame .n { left: 50%; top: 0; cursor: ns-resize; }
.picture-frame .ne { left: 100%; top: 0; cursor: nesw-resize; }
.picture-frame .e { left: 100%; top: 50%; cursor: ew-resize; }
.picture-frame .se { left: 100%; top: 100%; cursor: nwse-resize; }
.picture-frame .s { left: 50%; top: 100%; cursor: ns-resize; }
.picture-frame .sw { left: 0; top: 100%; cursor: nesw-resize; }
.picture-frame .w { left: 0; top: 50%; cursor: ew-resize; }
.picture-frame .rotate { left: 50%; top: -22px; cursor: grab; }
.picture-frame.cropping { border: 1px dashed #000; }
.picture-frame.cropping .handle { width: 10px; height: 10px; border-radius: 0; background: #000; border-color: #fff; }
.picture-panel details { border: 1px solid var(--border); border-radius: 4px; padding: 4px 8px; }
.picture-panel details[open] { display: flex; flex-direction: column; gap: 6px; padding-bottom: 8px; }
.picture-panel summary { cursor: pointer; font-weight: 600; }
.picture-panel .buttons { display: flex; flex-wrap: wrap; gap: 4px; }
.picture-panel .buttons button[aria-pressed="true"] { background: var(--active); }
.picture-panel input[type=color] { width: 100%; height: 26px; padding: 0 2px; }
.layout-options {
  position: absolute; z-index: 6; width: 26px; height: 26px; padding: 0; display: inline-flex; align-items: center; justify-content: center;
  border: 1px solid var(--border); border-radius: 4px; background: var(--surface); color: var(--text); cursor: pointer;
}
.picture-panel {
  position: fixed; z-index: 10; width: 300px; box-sizing: border-box; max-height: calc(100vh - 16px); overflow: auto; padding: 10px; display: flex; flex-direction: column; gap: 8px;
  background: var(--surface); color: var(--text); border: 1px solid var(--border); border-radius: 6px; box-shadow: 0 4px 16px rgb(0 0 0 / 15%); font-size: 13px;
}
.picture-panel fieldset { margin: 0; padding: 6px 8px; border: 1px solid var(--border); border-radius: 4px; display: flex; flex-direction: column; gap: 4px; }
.picture-panel fieldset:disabled { opacity: .5; }
.picture-panel legend { font-weight: 600; padding: 0 4px; }
.picture-panel .wraps { display: grid; grid-template-columns: 1fr 1fr; }
.picture-panel label { display: flex; flex-direction: column; gap: 2px; }
.picture-panel .wraps label, .picture-panel label.inline { flex-direction: row; align-items: center; gap: 4px; }
.picture-panel input[type=number], .picture-panel select, .picture-panel textarea { font: inherit; padding: 2px 4px; width: 100%; min-width: 0; box-sizing: border-box; }
.picture-panel .size-row { display: grid; grid-template-columns: 1fr 1fr; gap: 6px; }
.picture-panel button { font: inherit; padding: 3px 10px; border: 1px solid var(--border); border-radius: 4px; background: var(--surface); color: inherit; cursor: pointer; align-self: flex-start; }
.picture-panel .actions { display: flex; justify-content: flex-end; }
.change-menu button:disabled { opacity: .45; cursor: default; background: transparent; }
.split-dialog label { display: flex; justify-content: space-between; gap: 8px; align-items: center; }
.split-dialog input { width: 64px; }
.change-menu hr { width: 100%; margin: 4px 0; border: 0; border-top: 1px solid var(--border); }
.table-picker {
  position: fixed; z-index: 10; padding: 8px; display: flex; flex-direction: column; gap: 8px;
  background: var(--surface); border: 1px solid var(--border); border-radius: 6px; box-shadow: 0 4px 16px rgb(0 0 0 / 15%);
}
.table-picker .size { margin: 0; font-size: 13px; color: var(--muted); }
.table-picker .grid { display: grid; grid-template-columns: repeat(10, 18px); gap: 3px; }
.table-picker .cell { width: 18px; height: 18px; padding: 0; border: 1px solid var(--border); border-radius: 2px; background: var(--surface); cursor: pointer; }
.table-picker .cell.on { background: color-mix(in srgb, var(--accent) 25%, transparent); border-color: var(--accent); }
.table-picker .cell:focus-visible { outline: 2px solid var(--accent); outline-offset: 1px; }
.table-picker .custom { display: flex; gap: 6px; align-items: center; font-size: 13px; border-top: 1px solid var(--border); padding-top: 8px; }
.table-picker .custom label { display: flex; gap: 4px; align-items: center; }
.table-picker .custom input { width: 48px; font: inherit; padding: 2px 4px; }
.table-picker .custom button { font: inherit; padding: 2px 10px; border: 1px solid var(--border); border-radius: 4px; background: var(--surface); color: inherit; cursor: pointer; }
.palette {
  position: fixed; z-index: 10; width: 190px; padding: 8px; display: flex; flex-direction: column; gap: 8px;
  background: var(--surface); border: 1px solid var(--border); border-radius: 6px; box-shadow: 0 4px 16px rgb(0 0 0 / 15%);
}
.swatches { display: grid; grid-template-columns: repeat(5, 1fr); gap: 6px; }
.swatch { height: 24px; border: 1px solid rgb(0 0 0 / 25%); border-radius: 3px; cursor: pointer; padding: 0; }
.swatch.current, .swatch:focus-visible { outline: 2px solid var(--accent); outline-offset: 1px; }
.no-color { border: 1px solid var(--border); border-radius: 4px; background: var(--surface); color: inherit; padding: 4px; font: inherit; font-size: 13px; cursor: pointer; }
.marks-toggle { font-size: 16px !important; }
.track-toggle, .review-toggle { gap: 6px; padding: 0 8px !important; font-size: 13px !important; }
@media (max-width: 720px) { .track-toggle .label, .review-toggle .label { display: none; } }
.tracking-status { border: 0; background: none; color: #6f2dbd; font: inherit; padding: 0; cursor: pointer; text-decoration: underline dotted; }
.author-dialog { gap: 8px; padding: 12px; min-width: 260px; }
.author-dialog label { font-size: 13px; }
.author-dialog input { font: inherit; padding: 6px 8px; border: 1px solid var(--border); border-radius: 4px; }
.author-dialog .actions { display: flex; gap: 8px; justify-content: flex-end; }
.author-dialog .actions button { border: 1px solid var(--border); border-radius: 4px; padding: 5px 12px; background: var(--surface); color: inherit; cursor: pointer; text-align: center; }
.author-dialog .actions .confirm { background: var(--brand); border-color: var(--brand); color: #fff; }
.badge { font-size: 11px; color: var(--muted); }
.chooser {
  position: fixed; z-index: 10; min-width: 160px; max-width: 280px; max-height: 280px; overflow: auto;
  display: flex; flex-direction: column; padding: 4px; background: var(--surface); color: inherit;
  border: 1px solid var(--border); border-radius: 6px; box-shadow: 0 4px 16px rgb(0 0 0 / 15%);
}
.chooser button { text-align: left; padding: 6px 10px; border: 0; border-radius: 4px; background: transparent; color: inherit; font: inherit; cursor: pointer; }
.chooser button:hover, .chooser button:focus-visible, .chooser button[aria-selected="true"] { background: var(--hover); }
.chooser input { font: inherit; padding: 4px; }
.chooser .empty { margin: 6px 10px; color: var(--muted); }
.statusbar {
  display: flex; gap: 24px; padding: 4px 16px; font-size: 12px; color: var(--muted);
  background: var(--surface); border-top: 1px solid var(--border);
}

@media (max-width: 860px) {
  .canvas { padding: 0; }
  .docx-sheet { min-height: 100% !important; width: 100% !important; box-shadow: none; }
  .docx-page { padding: 24px 16px !important; }
  .docx-hf { padding: 12px 16px !important; min-height: 0 !important; }
  .docx-hf.header:has(.hf-tabs) { padding-top: 34px !important; }
  .docx-hf.footer:has(.hf-tabs) { padding-bottom: 34px !important; }
}
</style>

<style>
/* Unscoped: applies to the generated document HTML. */
.docx-sheet :is(h1, h2, h3, h4, h5, h6) { font-size: inherit; font-weight: inherit; }
.docx-sheet :is(p, h1, h2, h3, h4, h5, h6) { position: relative; }
.docx-sheet .cs-before { margin-top: 0 !important; }
.docx-sheet .cs-after { margin-bottom: 0 !important; }
.docx-sheet .list::before {
  content: attr(data-marker); position: absolute; transform: translateX(calc(-100% - 0.5em));
  text-indent: 0; white-space: nowrap;
}
.docx-sheet a { color: #0563c1; }
.docx-sheet img { max-width: 100%; vertical-align: bottom; }
/* Hidden markers (comment ranges, bookmarks, field codes) are empty: shown inline, not display:none,
   because the browser's editing drops unrendered nodes when it moves text, as in merging paragraphs. */
.docx-sheet .raw-hidden { display: none; }
.docx-sheet span.raw-hidden { display: inline; }
.docx-sheet .raw-block, .docx-sheet .raw-inline:not(:has(img)) {
  background: #f1f3f6; color: #555; border-radius: 3px; cursor: default; user-select: none;
}
.docx-sheet .raw-block { padding: 4px 8px; margin: 4px 0; border-left: 3px solid #c3c9d4; white-space: pre-wrap; }
.docx-sheet .raw-inline.deleted { text-decoration: line-through; color: #b3261e; background: transparent; }
.docx-sheet .raw-inline.inserted { text-decoration: underline; color: #1e6b34; background: transparent; }
/* Formatting marks (editor/marks.js): a layer over the page that never takes clicks. */
.docx-sheet .marks-layer { position: absolute; inset: 0; pointer-events: none; z-index: 1; overflow: hidden; }
.docx-sheet .fmark { position: absolute; color: #4f74b3; font-family: system-ui, sans-serif; white-space: pre; user-select: none; }
.docx-sheet .fmark.space { transform: translateX(-50%); }
/* Highlight, in Word's colours. */
.docx-sheet mark { background: #ffff00; color: inherit; }
.docx-sheet mark[data-color="green"] { background: #00ff00; } .docx-sheet mark[data-color="cyan"] { background: #00ffff; }
.docx-sheet mark[data-color="magenta"] { background: #ff00ff; } .docx-sheet mark[data-color="blue"] { background: #0000ff; }
.docx-sheet mark[data-color="red"] { background: #ff0000; } .docx-sheet mark[data-color="darkBlue"] { background: #000080; }
.docx-sheet mark[data-color="darkCyan"] { background: #008080; } .docx-sheet mark[data-color="darkGreen"] { background: #008000; }
.docx-sheet mark[data-color="darkMagenta"] { background: #800080; } .docx-sheet mark[data-color="darkRed"] { background: #800000; }
.docx-sheet mark[data-color="darkYellow"] { background: #808000; } .docx-sheet mark[data-color="darkGray"] { background: #808080; }
.docx-sheet mark[data-color="lightGray"] { background: #c0c0c0; } .docx-sheet mark[data-color="black"] { background: #000000; }
.docx-sheet mark[data-color="white"] { background: #ffffff; }
/* This session's tracked changes. */
.docx-sheet .tc-ins { color: #6f2dbd; text-decoration: underline; }
.docx-sheet .tc-del { color: #6f2dbd; text-decoration: line-through; cursor: default; }
.docx-sheet [data-tc-mark]::after { content: '¶'; color: #6f2dbd; margin-left: 1px; }
.docx-sheet [data-tc-mark="del"]::after { text-decoration: line-through; }
/* Tracked formatting changes (of a run: dashed underline; of a paragraph: a bar in the margin) and
   tracked table rows, as Word marks them. */
.docx-sheet [data-fmt="run"] { border-bottom: 1px dashed #8e44ad; }
.docx-sheet [data-fmt-para] { box-shadow: -10px 0 0 -8px #8e44ad; }
.docx-sheet table.docx-table[data-fmt-table] { outline: 2px dashed #8e5cc7; outline-offset: 2px; }
.docx-sheet tr[data-fmt-row] > td:first-child { box-shadow: inset 3px 0 0 #8e5cc7; }
.docx-sheet table.docx-table td[data-rev-merge] { box-shadow: inset 0 0 0 2px rgb(111 45 189 / 45%); }
.docx-sheet table.docx-table td[data-fmt-cell] { box-shadow: inset 0 0 0 1px #8e5cc7; }
.docx-sheet tr[data-rev-row="ins"] > td { background-color: #e8f5ea; }
.docx-sheet tr[data-rev-row="del"] > td { background-color: #fbe9e7; text-decoration: line-through; }
/* Someone else's tracked paragraph break, in the colours of their insertions and deletions. */
.docx-sheet [data-rev-mark]::after { content: '¶'; margin-left: 1px; }
.docx-sheet [data-rev-mark="ins"]::after { color: #1e6b34; }
.docx-sheet [data-rev-mark="del"]::after { color: #b3261e; text-decoration: line-through; }
.docx-sheet .raw-inline.picture { cursor: pointer; }
.docx-sheet .raw-inline.picture img { display: block; }
.docx-sheet .raw-inline.picture::selection, .docx-sheet .raw-inline.picture img::selection { background: transparent; } /* the frame shows it is selected */
.docx-sheet .raw-inline.picture[data-wrap="inline"] { display: inline-block; vertical-align: bottom; }
.docx-sheet .raw-inline.page-break { display: block; height: 0; border-top: 1px dashed #aab; margin: 8px 0; background: none; }
.docx-sheet .raw-inline.note-ref { vertical-align: super; font-size: smaller; background: transparent; color: #2b579a; }
.docx-sheet .raw-inline.note-ref::before { content: attr(data-n); }
.docx-sheet .raw-inline.symbol { background: transparent; color: inherit; }
.docx-sheet .raw-inline.comment-ref { display: inline; cursor: pointer; user-select: none; background: transparent; padding: 0; }
.docx-sheet .comment-ref::before {
  content: attr(data-n); display: inline-block; min-width: 1.3em; margin: 0 1px; padding: 0 3px; border-radius: 7px;
  font: 600 10px/1.4 system-ui, sans-serif; text-align: center; vertical-align: super; background: #d0705c; color: #fff;
}
.docx-sheet .field { border-bottom: 1px dotted #9aa6b8; }
::highlight(comment) { background-color: #fbe3dc; }
::highlight(field-near) { background-color: #d6e2f5; }
::highlight(comment-active) { background-color: #f6b9a8; }
.docx-sheet .sdt { border-radius: 2px; transition: background-color .1s, box-shadow .1s; }
.docx-sheet div.sdt { box-shadow: 0 0 0 1px transparent; }
.docx-sheet .sdt:not(.sdt-locked):hover, .docx-sheet .sdt:not(.sdt-locked):focus-within { background: #e8eef8; box-shadow: 0 0 0 1px #9db3d9; }
.docx-sheet .sdt-content { outline: none; }
.docx-sheet span.sdt-content:empty { padding: 0 1ch; }
.docx-sheet .showing-placeholder { color: #767676; }
.docx-sheet :is(.sdt-checkbox, .sdt-dropdown, .sdt-date) { cursor: pointer; }
.docx-sheet :is(span.sdt-dropdown, span.sdt-combo, span.sdt-date)::after { content: '▾'; font-size: .8em; color: #5b6b85; margin-left: 2px; cursor: pointer; }
.docx-sheet .sdt-locked { cursor: default; user-select: none; background: #f1f3f6; }
.docx-sheet table.docx-table { border-collapse: collapse; width: 100%; margin: 4px 0; }
.docx-sheet table.docx-table td { border: 1px solid #bfbfbf; padding: 2px 6px; vertical-align: top; outline: none; }
.docx-sheet table.docx-table td:focus-within { box-shadow: inset 0 0 0 2px #2b579a55; }
.docx-sheet table.docx-table td.vmerge { border-top-color: transparent; }
.docx-sheet table.docx-table td[data-rev-cell="ins"] { background-color: rgb(30 107 52 / 8%) !important; box-shadow: inset 0 0 0 1px rgb(30 107 52 / 45%); }
.docx-sheet table.docx-table td[data-rev-cell="del"] { background-color: rgb(179 38 30 / 8%) !important; box-shadow: inset 0 0 0 1px rgb(179 38 30 / 45%); }
.docx-sheet table.docx-table td.vmerge > * { visibility: hidden; }
</style>
