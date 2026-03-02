import pptx.util
import pptx.presentation
import pptx.slide
import pptx.shapes.picture
from lxml import etree

import pptx
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches


def xpath(el: pptx.oxml.shapes.ShapeElement, query: str):
    """Utility to query an `pptx.shapes.Shape`'s xml tree."""
    nsmap = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main"}
    return etree.ElementBase.xpath(el, query, namespaces=nsmap)


def autoplay_media(media: pptx.shapes.picture.Movie) -> None:
    """
    Utility to autoplay a media (currently only video) upon on entering the slide containing it.

    Args:
        media: `Shape` object containing the video.
    """
    el_id = xpath(media.element, ".//p:cNvPr")[0].attrib["id"]
    el_cnt = xpath(
        media.element.getparent().getparent().getparent(),
        './/p:timing//p:video//p:spTgt[@spid="%s"]' % el_id,
    )[0]
    cond = xpath(el_cnt.getparent().getparent(), ".//p:cond")[0]
    cond.set("delay", "0")
    cond.set("evt", "onBegin")


def move_slide(pres, from_index, to_index):
    """Move slide at position `from_index` in presentation `pres` to `to_index`"""
    slides = list(pres.slides._sldIdLst)
    if to_index < 0:
        to_index = len(slides) + to_index
    pres.slides._sldIdLst.remove(slides[from_index])
    pres.slides._sldIdLst.insert(to_index, slides[from_index])


def add_movie(
    pres: pptx.presentation.Presentation,
    slide: pptx.slide.Slide,
    movie_file: str,
    left: pptx.util.Length | int,
    top: pptx.util.Length | int,
    width: pptx.util.Length | int,
    height: pptx.util.Length | int,
    mime_type: str = "video/mp4",
    poster_frame_image: str | None = None,
    add_fullscreen: bool = True,
    hide_fullscreen_slide: bool = True,
) -> pptx.shapes.picture.Movie | tuple[pptx.shapes.picture.Movie | pptx.slide.Slide]:
    """
    Wrapper around add_movie method of a `pptx.slide.Slide` instance to add movies with functionality to toggle fullscreen mode

    Args:
        pres: the presentaion instance which contains the slide instance
        slide: slide instance to which we add the movie
        movie_file: path to the movie .mp4 file
        left: X-coordinate of the movie frame's top-left corner
        top: Y-coordinate of the movie frame's top-left corner
        width: width of the movie frame
        height: height of the movie frame
        mime_type: input to the mime_type keyword argument of slide.add_movie method.
        poster_frame_image: input to the poster_frame_image keyword argument of slide.add_movie method.
        add_fullscreen: Whether to add fullscreen toggling feature.
        hide_fullscreen_slide: Whether to hide the extra fullscreen slide or not. Recommend setting True if using PowerPoint and False if using Keynote.

    Returns:
        If `add_fullscreen == True`, returns a tuple of `(movie_shape, fullscreen_movie_slide)` else just returns the `movie_shape`
    """
    movie = slide.shapes.add_movie(
        "./example_video.mp4",
        left,
        top,
        width,
        height,
        mime_type="video/mp4",
        poster_frame_image=None,
    )
    if add_fullscreen:
        fs_btn_w, fs_btn_h = Inches(1.5), Inches(0.5)
        fs_btn_left = left + width - fs_btn_w - Inches(0.2)
        fs_btn_top = top + height - fs_btn_h - Inches(0.2)

        fs_btn = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, fs_btn_left, fs_btn_top, fs_btn_w, fs_btn_h
        )
        fs_btn.text = "Fullscreen"

        fs_movie_slide = pres.slides.add_slide(pres.slide_layouts[6])
        if hide_fullscreen_slide:
            fs_movie_slide.element.set("show", "0")

        fs_movie = fs_movie_slide.shapes.add_movie(
            "./example_video.mp4",
            0,
            0,
            pres.slide_width,
            pres.slide_height,
            mime_type="video/mp4",
            poster_frame_image=None,
        )
        autoplay_media(fs_movie)

        fs_btn.click_action.target_slide = fs_movie_slide

        fs_exit_btn = fs_movie_slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            pres.slide_width - fs_btn_w - Inches(0.2),
            Inches(0.2),
            fs_btn_h,
            fs_btn_h,
        )
        fs_exit_btn.text = "X"
        fs_exit_btn.click_action.target_slide = slide

        return movie, fs_movie_slide

    return movie
