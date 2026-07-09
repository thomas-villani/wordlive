"""CLI subcommands wired against the wordlive library."""

from __future__ import annotations

import click

from ..._ops import OP_REQUIRED_FIELDS as _OP_REQUIRED_FIELDS  # noqa: F401
from ..._ops import apply_op as _apply_op  # noqa: F401
from ..._ops import op_before as _op_before  # noqa: F401
from ..._ops import validate_op as _validate_op  # noqa: F401
from .charts import (
    add_error_bars_cmd,
    add_trendline_cmd,
    charts_cmd,
    format_axis_cmd,
    format_chart_cmd,
    format_series_cmd,
    insert_chart_cmd,
    set_series_color_cmd,
)
from .comments import comment
from .content_controls import create_content_control_cmd, set_cc_items_cmd, set_cc_properties_cmd
from .document import (
    checkpoint_cmd,
    cursor,
    diff_cmd,
    find_cmd,
    find_paragraph_cmd,
    go_to,
    locate_cmd,
    outline,
    paragraphs_cmd,
    pin_cmd,
    pin_outline_cmd,
    stats_cmd,
    status,
)
from .edit import append_cmd, delete_paragraph_cmd, prepend_cmd, replace, write
from .equations import equations_cmd, insert_equation_cmd
from .images import (
    images_cmd,
    read_image_cmd,
    set_image_alt_text_cmd,
    set_image_crop_cmd,
    set_image_size_cmd,
)
from .insert import (
    insert,
    insert_block_cmd,
    insert_break_cmd,
    insert_field_cmd,
    insert_image_cmd,
    insert_markdown_cmd,
    insert_section_cmd,
    insert_text_box_cmd,
    replace_section_cmd,
    update_fields_cmd,
)
from .linting import lint_cmd, proofing_cmd, regularize_cmd
from .lists import list_cmd
from .meta import exec_, install_mcp_cmd, install_skill_cmd, llm_help_cmd
from .metadata import properties, variables
from .persistence import export_pdf_cmd, save_as_cmd, save_cmd
from .read import read
from .references import (
    add_source_cmd,
    bibliography_style_cmd,
    bookmark,
    caption_cmd,
    cross_ref_cmd,
    endnotes_cmd,
    fields_cmd,
    footnotes_cmd,
    hyperlinks_cmd,
    insert_bibliography_cmd,
    insert_citation_cmd,
    insert_endnote_cmd,
    insert_footnote_cmd,
    insert_index_cmd,
    insert_toc_cmd,
    link_cmd,
    mark_citation_cmd,
    mark_index_entry_cmd,
    set_hyperlink_cmd,
    table_of_authorities_cmd,
    table_of_figures_cmd,
)
from .revisions import revision, revisions_cmd, track
from .sections import footer, header, page_setup_cmd, section, sections_cmd, watermark_cmd
from .shapes import (
    delete_shape_cmd,
    format_shape_cmd,
    group_shapes_cmd,
    replace_shape_image_cmd,
    set_shape_alt_text_cmd,
    set_shape_crop_cmd,
    set_shape_position_cmd,
    set_shape_rotation_cmd,
    set_shape_size_cmd,
    set_shape_text_cmd,
    set_shape_text_frame_cmd,
    set_shape_wrap_cmd,
    set_shape_z_order_cmd,
    shapes_cmd,
    ungroup_shape_cmd,
)
from .snapshot import snapshot_cmd
from .styles import (
    borders_cmd,
    drop_cap_cmd,
    format_paragraph_cmd,
    format_run_cmd,
    shading_cmd,
    style,
    tab_stop_cmd,
)
from .tables import cell_valign_cmd, table
from .theme import (
    apply_theme_cmd,
    list_themes_cmd,
    set_theme_colors_cmd,
    set_theme_fonts_cmd,
    theme_cmd,
)


def register(group: click.Group) -> None:
    group.add_command(status)
    group.add_command(outline)
    group.add_command(paragraphs_cmd)
    group.add_command(read)
    group.add_command(write)
    group.add_command(insert)
    group.add_command(insert_block_cmd)
    group.add_command(insert_section_cmd)
    group.add_command(insert_markdown_cmd)
    group.add_command(replace_section_cmd)
    group.add_command(delete_paragraph_cmd)
    group.add_command(insert_break_cmd)
    group.add_command(insert_field_cmd)
    group.add_command(update_fields_cmd)
    group.add_command(insert_footnote_cmd)
    group.add_command(insert_endnote_cmd)
    group.add_command(insert_toc_cmd)
    group.add_command(footnotes_cmd)
    group.add_command(endnotes_cmd)
    group.add_command(revisions_cmd)
    group.add_command(locate_cmd)
    group.add_command(stats_cmd)
    group.add_command(proofing_cmd)
    group.add_command(lint_cmd)
    group.add_command(regularize_cmd)
    group.add_command(checkpoint_cmd)
    group.add_command(diff_cmd)
    group.add_command(hyperlinks_cmd)
    group.add_command(set_hyperlink_cmd)
    group.add_command(fields_cmd)
    group.add_command(properties)
    group.add_command(variables)
    group.add_command(images_cmd)
    group.add_command(read_image_cmd)
    group.add_command(equations_cmd)
    group.add_command(insert_equation_cmd)
    group.add_command(charts_cmd)
    group.add_command(insert_chart_cmd)
    group.add_command(format_chart_cmd)
    group.add_command(format_axis_cmd)
    group.add_command(add_trendline_cmd)
    group.add_command(set_series_color_cmd)
    group.add_command(format_series_cmd)
    group.add_command(add_error_bars_cmd)
    group.add_command(shapes_cmd)
    group.add_command(set_shape_wrap_cmd)
    group.add_command(set_shape_crop_cmd)
    group.add_command(set_shape_position_cmd)
    group.add_command(set_shape_size_cmd)
    group.add_command(format_shape_cmd)
    group.add_command(set_shape_alt_text_cmd)
    group.add_command(set_shape_text_cmd)
    group.add_command(set_shape_rotation_cmd)
    group.add_command(set_shape_z_order_cmd)
    group.add_command(set_shape_text_frame_cmd)
    group.add_command(replace_shape_image_cmd)
    group.add_command(delete_shape_cmd)
    group.add_command(group_shapes_cmd)
    group.add_command(ungroup_shape_cmd)
    group.add_command(set_image_alt_text_cmd)
    group.add_command(set_image_size_cmd)
    group.add_command(set_image_crop_cmd)
    group.add_command(bookmark)
    group.add_command(pin_cmd)
    group.add_command(pin_outline_cmd)
    group.add_command(link_cmd)
    group.add_command(cross_ref_cmd)
    group.add_command(caption_cmd)
    group.add_command(create_content_control_cmd)
    group.add_command(set_cc_properties_cmd)
    group.add_command(set_cc_items_cmd)
    group.add_command(mark_index_entry_cmd)
    group.add_command(insert_index_cmd)
    group.add_command(table_of_figures_cmd)
    group.add_command(bibliography_style_cmd)
    group.add_command(add_source_cmd)
    group.add_command(insert_citation_cmd)
    group.add_command(insert_bibliography_cmd)
    group.add_command(mark_citation_cmd)
    group.add_command(table_of_authorities_cmd)
    group.add_command(theme_cmd)
    group.add_command(list_themes_cmd)
    group.add_command(apply_theme_cmd)
    group.add_command(set_theme_colors_cmd)
    group.add_command(set_theme_fonts_cmd)
    group.add_command(page_setup_cmd)
    group.add_command(prepend_cmd)
    group.add_command(append_cmd)
    group.add_command(insert_image_cmd)
    group.add_command(snapshot_cmd)
    group.add_command(save_cmd)
    group.add_command(save_as_cmd)
    group.add_command(export_pdf_cmd)
    group.add_command(cursor)
    group.add_command(find_cmd)
    group.add_command(find_paragraph_cmd)
    group.add_command(replace)
    group.add_command(go_to)
    group.add_command(style)
    group.add_command(format_paragraph_cmd)
    group.add_command(format_run_cmd)
    group.add_command(shading_cmd)
    group.add_command(borders_cmd)
    group.add_command(cell_valign_cmd)
    group.add_command(drop_cap_cmd)
    group.add_command(tab_stop_cmd)
    group.add_command(table)
    group.add_command(comment)
    group.add_command(track)
    group.add_command(revision)
    group.add_command(watermark_cmd)
    group.add_command(insert_text_box_cmd)
    group.add_command(list_cmd)
    group.add_command(sections_cmd)
    group.add_command(section)
    group.add_command(header)
    group.add_command(footer)
    group.add_command(exec_)
    group.add_command(llm_help_cmd)
    group.add_command(install_skill_cmd)
    group.add_command(install_mcp_cmd)


# ---------------------------------------------------------------------------
# status
# ---------------------------------------------------------------------------
