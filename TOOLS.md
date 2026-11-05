# Tools

visio-mcp has 25 tools. This page is generated from `tests/snapshots/live_tools.json`
by `scripts/tools_md.py`, so it matches what an MCP client is sent.

Positions and sizes are in inches, with x and y the shape's centre measured from the page's
bottom-left corner. A shape is named by its ID, else its exact name, else its exact text.
Every tool that changes a drawing is one Ctrl+Z step, and rolls back if it fails part-way.

| Group | Tools |
|---|---|
| Generate | `visio_create_diagram`, `visio_find_masters`, `visio_list_templates`, `visio_layout_page` |
| Edit | `visio_add_elements`, `visio_update_shapes`, `visio_delete_shapes`, `visio_add_container`, `visio_add_page`, `visio_update_page`, `visio_delete_page`, `visio_undo`, `visio_redo` |
| Read | `visio_list_open`, `visio_get_page`, `visio_get_shape`, `visio_find_shapes`, `visio_get_selection`, `visio_get_page_image`, `visio_diagnose_page` |
| Files | `visio_new_drawing`, `visio_open_drawing`, `visio_save`, `visio_export`, `visio_close_drawing` |

## Generate

### visio_create_diagram

**Create Visio Diagram.** Changes the drawing.

Build a whole diagram in a new drawing in one call: shapes, glued connectors, layout and a picture.

This is the way to make a diagram. Find masters with visio_find_masters first, then pass
them all at once. It starts a visible Visio if none is running. A bad spec changes nothing,
and one Ctrl+Z removes everything it adds. To add to an existing drawing, use visio_add_elements.

Returns the drawing, its page, ids mapping your keys to shape IDs, the connectors, and
issues found on the page, with a PNG to check the result by eye.

Parameters:

- `spec` (object, required)
  - `nodes` (list of object, required)
    - `key` (str, required): Your name for this node, which edges and the returned ids use.
    - `master` (str, required): A master like 'BASFLO:Decision'. Find one with visio_find_masters.
    - `text` (str)
    - `data` (map of str to str): Shape Data fields, by label.
    - `x` (number): Centre, in inches from the page's left. Only with layout style 'none'.
    - `y` (number): Centre, in inches from the page's bottom. Only with layout style 'none'.
    - `w` (number): Width in inches. Leave out to use the master's.
    - `h` (number): Height in inches. Setting it stops the shape growing to fit its text.
    - `fill` (str): Fill colour as '#RRGGBB'.
  - `edges` (list of object)
    - `from` (str, required): The key of the node it starts at. In visio_add_elements, also an existing shape's ID, name or text.
    - `to` (str, required): The key of the node it ends at, or as for from.
    - `text` (str)
    - `arrow` ("end" or "both" or "none", default "end")
  - `layout` (object): Defaults to top_down, or to none if the nodes have x and y.
    - `style` ("top_down" or "left_right" or "radial" or "circular" or "none", default "top_down"): top_down and left_right give the cleanest results. radial and circular are weaker. none keeps each node's x and y.
    - `route` ("flowchart" or "tree" or "right_angle" or "straight"): Connector style. Defaults to flowchart, or straight for radial and circular.
    - `spacing` (number): Gap between shapes in inches. Visio's default is 0.5.
    - `fit_page` (bool): Resize the page to fit its shapes. Defaults to true, or false for style 'none', since it moves shapes.
  - `template` (str): A name from visio_list_templates, or a template file's absolute path.
  - `units` ("us" or "metric")
  - `stencils` (list of str): Stencils to search for masters named without one, like 'BASFLO'.
  - `save_as` (str): An absolute path to save to.
  - `overwrite` (bool, default false): Replace the file at save_as if there is one.
  - `return_image` (bool, default true)

### visio_find_masters

**Find Visio Masters.** Reads only.

Search Visio's installed shapes for masters to use in a diagram.

Matches master names first, then stencil titles and keywords, so "flowchart decision"
finds BASFLO:Decision. Pass stencil alone to list one stencil's masters. Use the returned
"master" value in visio_create_diagram and visio_add_elements. Does not need Visio running.

Parameters:

- `query` (str, default "")
- `stencil` (str)
- `limit` (int, default 25)

### visio_list_templates

**List Visio Templates.** Reads only.

List Visio's installed templates, optionally filtered by words in their title or keywords.

Pass the returned "template" value to visio_new_drawing or visio_create_diagram.
Does not need Visio running.

Parameters:

- `query` (str)

### visio_layout_page

**Lay Out Visio Page.** Changes the drawing.

Lay out a page again, or only some of its shapes. Connectors stay glued, and one Ctrl+Z undoes it.

top_down and left_right give the cleanest results. radial and circular are weaker. route
defaults to flowchart, or straight for radial and circular. spacing is the gap in inches.
shapes limits it to those shapes, as IDs, names or exact text. fit resizes the page to its
shapes, by default only when laying out the whole page. Shapes with no connections stay where
they are unless they're marked placeable, so check the returned issues for overlaps.

Parameters:

- `style` ("top_down" or "left_right" or "radial" or "circular", default "top_down")
- `route` ("flowchart" or "tree" or "right_angle" or "straight")
- `spacing` (number)
- `shapes` (list of int or str)
- `fit` (bool)
- `drawing` (str)
- `page` (int or str)

## Edit

### visio_add_elements

**Add Visio Elements.** Changes the drawing.

Add shapes and connectors to an existing page, as one Ctrl+Z step.

nodes are as in visio_create_diagram. An edge's from and to name a new node's key, else an
existing shape's ID, name or exact text, so new nodes can join what's already there. Keys
win over existing shapes.

Placing: with x and y on every node, nothing else moves. Without them, the whole page is laid
out again in the style it was last laid out with (else top_down), which moves existing shapes
too. Edges alone move nothing unless layout is passed.

Returns ids mapping your keys to shape IDs, the connectors, and issues found on the page.
Look at the result with visio_get_page_image.

Parameters:

- `nodes` (list of object)
  - `key` (str, required): Your name for this node, which edges and the returned ids use.
  - `master` (str, required): A master like 'BASFLO:Decision'. Find one with visio_find_masters.
  - `text` (str)
  - `data` (map of str to str): Shape Data fields, by label.
  - `x` (number): Centre, in inches from the page's left. Only with layout style 'none'.
  - `y` (number): Centre, in inches from the page's bottom. Only with layout style 'none'.
  - `w` (number): Width in inches. Leave out to use the master's.
  - `h` (number): Height in inches. Setting it stops the shape growing to fit its text.
  - `fill` (str): Fill colour as '#RRGGBB'.
- `edges` (list of object)
  - `from` (str, required): The key of the node it starts at. In visio_add_elements, also an existing shape's ID, name or text.
  - `to` (str, required): The key of the node it ends at, or as for from.
  - `text` (str)
  - `arrow` ("end" or "both" or "none", default "end")
- `layout` (object)
  - `style` ("top_down" or "left_right" or "radial" or "circular" or "none", default "top_down"): top_down and left_right give the cleanest results. radial and circular are weaker. none keeps each node's x and y.
  - `route` ("flowchart" or "tree" or "right_angle" or "straight"): Connector style. Defaults to flowchart, or straight for radial and circular.
  - `spacing` (number): Gap between shapes in inches. Visio's default is 0.5.
  - `fit_page` (bool): Resize the page to fit its shapes. Defaults to true, or false for style 'none', since it moves shapes.
- `drawing` (str)
- `page` (int or str)

### visio_update_shapes

**Update Visio Shapes.** Changes the drawing.

Change shapes' text, Shape Data, position, size or fill, as one Ctrl+Z step.

Each update names a shape by ID, else exact name, else exact text, and sets only the fields
given. x and y are the centre in inches from the page's bottom-left, and glued connectors
follow a moved shape. Every shape is found before anything changes.

Returns each shape's ID, name, text and master afterwards. Check the ID and text are the ones
you meant, since Visio can reuse an ID after a delete.

Parameters:

- `updates` (list of object, required)
  - `shape` (int or str, required): The shape's ID, else its exact name, else its exact text.
  - `text` (str)
  - `data` (map of str to str): Shape Data fields to set, by label. A new label adds a field, and an empty value clears one.
  - `x` (number): New centre, in inches from the page's left.
  - `y` (number): New centre, in inches from the page's bottom.
  - `w` (number): Width in inches.
  - `h` (number): Height in inches. Setting it stops the shape growing to fit its text.
  - `fill` (str): Fill colour as '#RRGGBB'.
- `drawing` (str)
- `page` (int or str)

### visio_delete_shapes

**Delete Visio Shapes.** May delete or overwrite something.

Delete shapes, and the connectors glued to them, as one Ctrl+Z step.

shapes are IDs, else exact names, else exact text. Deleting a container keeps the shapes in
it, and deleting a group deletes the group's own members. Returns what was deleted, with the
connectors that went with it listed apart.

Parameters:

- `shapes` (list of int or str, required)
- `drawing` (str)
- `page` (int or str)

### visio_add_container

**Add Visio Container.** Changes the drawing.

Put a titled container around shapes, as one Ctrl+Z step.

shapes are IDs, else exact names, else exact text. The container is sized to fit them, and
moving it in Visio moves them too. master defaults to 'SDCONT:Plain'; find other styles with
visio_find_masters("container"). Returns the container and the shape IDs it holds.

Parameters:

- `shapes` (list of int or str, required)
- `title` (str, required)
- `master` (str)
- `drawing` (str)
- `page` (int or str)

### visio_add_page

**Add Visio Page.** Changes the drawing.

Add a page to a drawing, as one Ctrl+Z step. Pass its name as page in later calls.

Parameters:

- `name` (str)
- `drawing` (str)

### visio_update_page

**Update Visio Page.** Changes the drawing.

Rename a page, set its size or theme, or fit it to its shapes, as one Ctrl+Z step.

width and height are in inches. fit sets the size from the shapes, so it can't go with them.
theme is a name such as Office, Linear, Simple, Facet, Ion, Slice, Retrospect, Whisp, Organic,
Pencil or Whiteboard, ignoring case, or an index from Visio's gallery. 'None' or 0 removes it.
Returns the page with its theme index, and issues found on it.

Parameters:

- `page` (int or str)
- `name` (str)
- `width` (number)
- `height` (number)
- `fit` (bool)
- `theme` (str or int)
- `drawing` (str)

### visio_delete_page

**Delete Visio Page.** May delete or overwrite something.

Delete a page and everything on it, as one Ctrl+Z step.

page is a name or a 1-based number. A drawing's only foreground page is refused, because
Visio would wipe it and make a blank one.

Parameters:

- `page` (int or str, required)
- `drawing` (str)

### visio_undo

**Undo in Visio.** Changes the drawing.

Undo the last steps in Visio, like pressing Ctrl+Z. Each edit tool call is one step.

Visio keeps one history for all its drawings, and the user's own edits are in it, so a step
can land in another drawing or take back something the user did. The shape and page counts
per drawing, before and after, show where it landed. They don't show text or position
changes, so check with visio_get_page. Steps beyond the start of the history do nothing.

Parameters:

- `steps` (int, default 1)

### visio_redo

**Redo in Visio.** Changes the drawing.

Redo steps undone in Visio, like pressing Ctrl+Y.

Visio keeps one history for all its drawings, so a step can land in another drawing. The
shape and page counts per drawing, before and after, show where it landed. They don't show
text or position changes, so check with visio_get_page. Steps with nothing to redo do nothing.

Parameters:

- `steps` (int, default 1)

## Read

### visio_list_open

**List Open Visio Drawings.** Reads only.

List the drawings open in Visio, saying which is active. Never starts Visio.

Reports running=false if Visio isn't running. Includes saved drawings in other Visio
windows where Windows lets them be reached.

No parameters.

### visio_get_page

**Read Visio Page.** Reads only.

List a page's top-level shapes and connectors, one page of results at a time.

Shapes have centre x, y and size w, h in inches, from the bottom-left of the page.
Connectors say which shape IDs they join in "from" and "to", where null means not glued.
Containers list the shape IDs they hold in "contains".
Pass next_offset back as offset to continue. Use visio_get_shape for a shape's data.

Parameters:

- `offset` (int, default 0)
- `limit` (int, default 100)
- `drawing` (str)
- `page` (int or str)

### visio_get_shape

**Read Visio Shape.** Reads only.

Read one shape in detail: position, Shape Data, connections, containers it's in, what it holds.

shape is an ID, else an exact name, else exact text. cells adds ShapeSheet cells by name,
such as "FillForegnd", with null for a cell the shape doesn't have.

Parameters:

- `shape` (int or str, required)
- `cells` (list of str)
- `drawing` (str)
- `page` (int or str)

### visio_find_shapes

**Find Visio Shapes.** Reads only.

Find a page's top-level shapes by text (part of it, ignoring case), master name, or Shape Data.

Filters combine. data matches a field's label or row name against an exact value,
ignoring case. Returns at most 200 shapes with the total found.

Parameters:

- `text` (str)
- `master` (str)
- `data` (map of str to str)
- `drawing` (str)
- `page` (int or str)

### visio_get_selection

**Read Visio Selection.** Reads only.

Read what's selected in the active Visio window, so you can act on what the user points at.

No parameters.

### visio_get_page_image

**Get Visio Page Image.** Reads only.

Get a PNG of a page, to check how a diagram looks. Doesn't change focus or the drawing.

The image is cropped to the shapes, and its longest side is at most max_px, unless a very
large page needs more to stay legible. annotate_ids badges each shape with its ID.

Parameters:

- `max_px` (int, default 1400)
- `annotate_ids` (bool, default false)
- `drawing` (str)
- `page` (int or str)

### visio_diagnose_page

**Diagnose Visio Page.** Reads only.

Check a page for overlapping shapes, shapes off the page, and connectors not glued at both ends.

Overlaps compare shapes' boxes, so a container and its own members don't count.
Connectors crossing shapes, and text too big for its shape, aren't checked, so look at
visio_get_page_image as well. Lists at most 50 issues, with issue_count saying if there are more.

Parameters:

- `drawing` (str)
- `page` (int or str)

## Files

### visio_new_drawing

**New Visio Drawing.** Changes the drawing.

Create a new drawing, starting a visible Visio if none is running.

template is a name from visio_list_templates or the absolute path of a template file.
Omit it for a blank drawing. units picks US or metric paper and shapes.

Parameters:

- `template` (str)
- `units` ("us" or "metric")

### visio_open_drawing

**Open Visio Drawing.** Changes the drawing.

Open a saved drawing, starting a visible Visio if none is running.

If it's already open, returns that one with already_open=true, even if it was opened
with different access. Check read_only in the result.

Parameters:

- `path` (str, required)
- `read_only` (bool, default false)

### visio_save

**Save Visio Drawing.** May delete or overwrite something.

Save a drawing, or save a copy to a new absolute path and continue in that copy.

Refuses to replace another file unless overwrite is true. A read-only drawing can only
be saved to a new path, and the result is writable.

Parameters:

- `path` (str)
- `overwrite` (bool, default false)
- `drawing` (str)

### visio_export

**Export Visio Drawing.** May delete or overwrite something.

Export to an absolute path as PNG, JPG, SVG or PDF. The format comes from the extension.

Images export one page: the named one, else the one on screen. PDF exports the named
page, else every page. Refuses to replace a file unless overwrite is true.

Parameters:

- `path` (str, required)
- `format` ("png" or "jpg" or "svg" or "pdf")
- `overwrite` (bool, default false)
- `drawing` (str)
- `page` (int or str)

### visio_close_drawing

**Close Visio Drawing.** May delete or overwrite something.

Close a drawing. Visio itself stays open.

Refuses if there are unsaved changes unless discard is true, which loses them and
must name the drawing, so it can't take whichever drawing happens to be active.

Parameters:

- `discard` (bool, default false)
- `drawing` (str)
