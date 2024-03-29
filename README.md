# VisioStyleResources
Visio Style Resources was created to assist in projects that mimic or expand on Visio's user-interface. When creating custom forms, it is sometimes necessary to re-create visuals, such as arrowheads, fill patterns, line patterns, and other stylistic entities.

It is the goal of this repository to provide icons and images that will make this less of a chore for Visio add-in developers. 

For example, here are some of the built-in Visio arrowhead styles, exported as 128 x 128 pixel icons (pardon the DropBox adornments in Explorer...):

![Sample Arrowhead Icons](https://github.com/visioguy/VisioStyleResources/blob/master/img/sample_arrowhead_icons.png)

And the line patterns:

![Sample Line Pattern Icons](https://github.com/visioguy/VisioStyleResources/blob/master/img/sample_line_pattern_icons.png)

Currently, VisioStyleResources has icons for:

- Icons for all 46 (0-45) of Visio's built-in arrowhead styles in 32x23, 64x64, and 128x128 sets.
- Icons for all 23 (1-23) of Visio's built-in line patterns, in 128x32, 256x64, 512x128 sets.

The icon sets also include single "filmstrip" images (_allIcons_xx.png) that can be frame-shifted in place of multiple image files.

In the future:
- Upload macro-enabled Visio drawing that is used to automate the creation of the resources.
- Upload the exported VBA code from that document's VBA project.
- Visio fill patterns?

## Note About Code in Directory 'src'
The icons were generate from within Visio, using a combination of Visio SmartShapes and VBA (Visual Basic) code. I've uploaded the code from that document, because some of the procedures and techniques might be useful for folks needing to export Visio shapes. But the code realy makes 100% sense only if the Visio file is also present, since the SmartShapes in that document call the code, and the code depends on certain features of those shapes.

Once the Visio file is sufficiently cleaned up, I will add it to the repository.
