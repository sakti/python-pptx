# -*- coding: utf-8 -*-
#
# presentation.py
#
# Copyright (C) 2012, 2013 Steve Canny scanny@cisco.com
#
# This module is part of python-pptx and is released under
# the MIT License: http://www.opensource.org/licenses/mit-license.php

"""
API classes for dealing with presentations and other objects one typically
encounters as an end-user of the PowerPoint user interface.
"""

import hashlib

try:
    from PIL import Image as PIL_Image
except ImportError:
    import Image as PIL_Image

import os
import posixpath
import weakref

import pptx.packaging
import pptx.spec as spec
import pptx.util as util

from pptx.exceptions import InvalidPackageError
from pptx.oxml import (
    _Element, _SubElement, oxml_fromstring, oxml_tostring, qn)

from pptx.shapes import ShapeCollection
from pptx.spec import namespaces
from pptx.spec import (
    CT_PRESENTATION, CT_SLIDE, CT_SLIDE_LAYOUT, CT_SLIDE_MASTER, CT_SLIDESHOW,
    CT_TEMPLATE)
from pptx.spec import (
    RT_IMAGE, RT_OFFICE_DOCUMENT, RT_SLIDE, RT_SLIDE_LAYOUT, RT_SLIDE_MASTER)

from pptx.util import Collection

import logging
log = logging.getLogger('pptx.presentation')
log.setLevel(logging.DEBUG)
# log.setLevel(logging.INFO)
ch = logging.StreamHandler()
ch.setLevel(logging.DEBUG)
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - '
                              '%(message)s')
ch.setFormatter(formatter)
log.addHandler(ch)

# default namespace map for use in lxml calls
_nsmap = namespaces('a', 'r', 'p')


def _child(element, child_tagname, nsmap=None):
    """
    Return direct child of *element* having *child_tagname* or :class:`None`
    if no such child element is present.
    """
    # use default nsmap if not specified
    if nsmap is None:
        nsmap = _nsmap
    xpath = './%s' % child_tagname
    matching_children = element.xpath(xpath, namespaces=nsmap)
    return matching_children[0] if len(matching_children) else None


def _to_unicode(text):
    """
    Return *text* as a unicode string.

    *text* can be a 7-bit ASCII string, a UTF-8 encoded 8-bit string, or
    unicode. String values are converted to unicode assuming UTF-8 encoding.
    Unicode values are returned unchanged.
    """
    # both str and unicode inherit from basestring
    if not isinstance(text, basestring):
        tmpl = 'expected UTF-8 encoded string or unicode, got %s value %s'
        raise TypeError(tmpl % (type(text), text))
    # return unicode strings unchanged
    if isinstance(text, unicode):
        return text
    # otherwise assume UTF-8 encoding, which also works for ASCII
    return unicode(text, 'utf-8')


# ============================================================================
# Package
# ============================================================================

class Package(object):
    """
    Return an instance of |Package| loaded from *file*, where *file* can be a
    path (a string) or a file-like object. If *file* is a path, it can be
    either a path to a PowerPoint `.pptx` file or a path to a directory
    containing an expanded presentation file, as would result from unzipping
    a `.pptx` file. If *file* is |None|, the default presentation template is
    loaded.
    """
    # track instances as weakrefs so .containing() can be computed
    __instances = []

    def __init__(self, file=None):
        super(Package, self).__init__()
        self.__presentation = None
        self.__relationships = _RelationshipCollection()
        self.__images = ImageCollection()
        self.__instances.append(weakref.ref(self))
        if file is None:
            file = self.__default_pptx_path
        self.__open(file)

    @classmethod
    def containing(cls, part):
        """Return package instance that contains *part*"""
        for pkg in cls.instances():
            if part in pkg._parts:
                return pkg
        raise KeyError("No package contains part %r" % part)

    @classmethod
    def instances(cls):
        """Return tuple of Package instances that have been created"""
        # clean garbage collected pkgs out of __instances
        cls.__instances[:] = [wkref for wkref in cls.__instances
                              if wkref() is not None]
        # return instance references in a tuple
        pkgs = [wkref() for wkref in cls.__instances]
        return tuple(pkgs)

    @property
    def presentation(self):
        """
        Reference to the :class:`Presentation` instance contained in this
        package.
        """
        return self.__presentation

    def save(self, file):
        """
        Save this package to *file*, where *file* can be either a path to a
        file (a string) or a file-like object.
        """
        pkgng_pkg = pptx.packaging.Package().marshal(self)
        pkgng_pkg.save(file)

    @property
    def _images(self):
        return self.__images

    @property
    def _relationships(self):
        return self.__relationships

    def __load(self, pkgrels):
        """
        Load all the model-side parts and relationships from the on-disk
        package by loading package-level relationship parts and propagating
        the load down the relationship graph.
        """
        # keep track of which parts are already loaded
        part_dict = {}

        # discard any previously loaded relationships
        self.__relationships = _RelationshipCollection()

        # add model-side rel for each pkg-side one, and load target parts
        for pkgrel in pkgrels:
            # unpack working values for part to be loaded
            reltype = pkgrel.reltype
            pkgpart = pkgrel.target
            partname = pkgpart.partname
            content_type = pkgpart.content_type
            # log.debug("%s -- %s", reltype, partname)

            # create target part
            part = Part(reltype, content_type)
            part_dict[partname] = part
            part._load(pkgpart, part_dict)

            # create model-side package relationship
            model_rel = _Relationship(pkgrel.rId, reltype, part)
            self.__relationships._additem(model_rel)

        # gather references to image parts into __images
        self.__images = ImageCollection()
        image_parts = [part for part in self._parts
                       if part.__class__.__name__ == 'Image']
        for image in image_parts:
            self.__images._loadpart(image)

    def __open(self, file):
        """
        Load presentation contained in *file* into this package.
        """
        pkg = pptx.packaging.Package().open(file)
        self.__load(pkg.relationships)
        # unmarshal relationships selectively for now
        for rel in self.__relationships:
            if rel._reltype == RT_OFFICE_DOCUMENT:
                self.__presentation = rel._target

    @property
    def __default_pptx_path(self):
        """
        The path of the default presentation, used when no path is specified
        on construction.
        """
        thisdir = os.path.split(__file__)[0]
        return os.path.join(thisdir, 'templates', 'default.pptx')

    @property
    def _parts(self):
        """
        Return a list containing a reference to each of the parts in this
        package.
        """
        return [part for part in Package.__walkparts(self.__relationships)]

    @staticmethod
    def __walkparts(rels, parts=None):
        """
        Recursive function, walk relationships to iterate over all parts in
        this package. Leave out *parts* parameter in call to visit all parts.
        """
        # initial call can leave out parts parameter as a signal to initialize
        if parts is None:
            parts = []
        # log.debug("in __walkparts(), len(parts)==%d", len(parts))
        for rel in rels:
            # log.debug("rel.target.partname==%s", rel.target.partname)
            part = rel._target
            # only visit each part once (graph is cyclic)
            if part in parts:
                continue
            parts.append(part)
            yield part
            for part in Package.__walkparts(part._relationships, parts):
                yield part


# ============================================================================
# Base classes
# ============================================================================

class Observable(object):
    """
    Simple observer pattern mixin. Limitations:

    * observers get all message types from subject (Observable), subscription
      is on subject basis, not subject + event_type.

    * notifications are oriented toward "value has been updated", which seems
      like only one possible event, could also be something like "load has
      completed" or "preparing to load".
    """
    def __init__(self):
        super(Observable, self).__init__()
        self._observers = []

    def _notify_observers(self, name, value):
        # value = getattr(self, name)
        for observer in self._observers:
            observer.notify(self, name, value)

    def add_observer(self, observer):
        """
        Begin notifying *observer* of events. *observer* must implement method
        ``notify(observed, name, new_value)``
        """
        if observer not in self._observers:
            self._observers.append(observer)

    def remove_observer(self, observer):
        """Remove *observer* from notification list."""
        assert observer in self._observers, "remove_observer called for"\
                                            "unsubscribed object"
        self._observers.remove(observer)


# ============================================================================
# Relationships
# ============================================================================

class _RelationshipCollection(Collection):
    """
    Sequence of relationships maintained in rId order. Maintaining the
    relationships in sorted order makes the .rels files both repeatable and
    more readable, which is very helpful for debugging.
    :class:`RelationshipCollection` has an attribute *_reltype_ordering* which
    is a sequence (tuple) of reltypes. If *_reltype_ordering* contains one or
    more reltype, the collection is maintained in reltype + partname.idx
    order and relationship ids (rIds) are renumbered to match that sequence
    and any numbering gaps are filled in.
    """
    def __init__(self):
        super(_RelationshipCollection, self).__init__()
        self.__reltype_ordering = ()

    def _additem(self, relationship):
        """
        Insert *relationship* into the appropriate position in this ordered
        collection.
        """
        rIds = [rel._rId for rel in self._values]
        if relationship._rId in rIds:
            tmpl = "cannot add relationship with duplicate rId '%s'"
            raise ValueError(tmpl % relationship._rId)
        self._values.append(relationship)
        self.__resequence()
        # register as observer of partname changes
        relationship._target.add_observer(self)

    @property
    def _next_rId(self):
        """
        Next available rId in collection, starting from 'rId1' and making use
        of any gaps in numbering, e.g. 'rId2' for rIds ['rId1', 'rId3'].
        """
        tmpl = 'rId%d'
        next_rId_num = 1
        for relationship in self._values:
            if relationship._num > next_rId_num:
                return tmpl % next_rId_num
            next_rId_num += 1
        return tmpl % next_rId_num

    @property
    def _reltype_ordering(self):
        """
        Tuple of relationship types, e.g. ``(RT_SLIDE, RT_SLIDE_LAYOUT)``. If
        present, relationships of those types are grouped, and those groups
        are ordered in the same sequence they appear in the tuple. In
        addition, relationships of the same type are sequenced in order of
        partname number (e.g. 16 for /ppt/slides/slide16.xml). If empty, the
        collection is maintained in existing rId order; rIds are not
        renumbered and any gaps in numbering are left to remain. Specifying
        this value for a collection with members causes it to be immediately
        reordered. The ordering is maintained as relationships are added or
        removed, renumbering rIds whenever necessary to also maintain the
        sequence in rId order.
        """
        return self.__reltype_ordering

    @_reltype_ordering.setter
    def _reltype_ordering(self, ordering):
        self.__reltype_ordering = tuple(ordering)
        self.__resequence()

    def rels_of_reltype(self, reltype):
        """
        Return a :class:`list` containing the subset of relationships in this
        collection of type *reltype*. The returned list is ordered by rId.
        Returns an empty list if there are no relationships of type *reltype*
        in the collection.
        """
        return [rel for rel in self._values if rel._reltype == reltype]

    def notify(self, subject, name, value):
        """RelationshipCollection implements the Observer interface"""
        if isinstance(subject, BasePart):
            if name == 'partname':
                self.__resequence()

    def __resequence(self):
        """
        Sort relationships and renumber if necessary to maintain values in rId
        order.
        """
        if self.__reltype_ordering:
            def reltype_key(rel):
                reltype = rel._reltype
                if reltype in self.__reltype_ordering:
                    return self.__reltype_ordering.index(reltype)
                return len(self.__reltype_ordering)

            def partname_idx_key(rel):
                partname = util.Partname(rel._target.partname)
                if partname.idx is None:
                    return 0
                return partname.idx
            self._values.sort(key=lambda rel: partname_idx_key(rel))
            self._values.sort(key=lambda rel: reltype_key(rel))
            # renumber consistent with new sort order
            for idx, relationship in enumerate(self._values):
                relationship._rId = 'rId%d' % (idx+1)
        else:
            self._values.sort(key=lambda rel: rel._num)


class _Relationship(object):
    """
    Relationship to a part from a package or part. *rId* must be unique in any
    |_RelationshipCollection| this relationship is added to; use
    :attr:`_RelationshipCollection._next_rId` to get a unique rId.
    """
    def __init__(self, rId, reltype, target):
        super(_Relationship, self).__init__()
        # can't get _num right if rId is non-standard form
        assert rId.startswith('rId'), "rId in non-standard form: '%s'" % rId
        self.__rId = rId
        self.__reltype = reltype
        self.__target = target

    @property
    def _rId(self):
        """
        Relationship id for this relationship. Must be of the form
        ``rId[1-9][0-9]*``.
        """
        return self.__rId

    @property
    def _reltype(self):
        """
        Relationship type URI for this relationship. Corresponds roughly to
        the part type of the target part.
        """
        return self.__reltype

    @property
    def _target(self):
        """
        Target part of this relationship. Relationships are directed, from a
        source and a target. The target is always a part.
        """
        return self.__target

    @property
    def _num(self):
        """
        The numeric portion of the rId of this |_Relationship|, expressed as
        an :class:`int`. For example, :attr:`_num` for a relationship with an
        rId of ``'rId12'`` would be ``12``.
        """
        try:
            num = int(self.__rId[3:])
        except ValueError:
            num = 9999
        return num

    @_rId.setter
    def _rId(self, value):
        self.__rId = value


# ============================================================================
# Parts
# ============================================================================

class PartCollection(Collection):
    """
    Sequence of parts. Sensitive to partname index when ordering parts added
    via _loadpart(), e.g. ``/ppt/slide/slide2.xml`` appears before
    ``/ppt/slide/slide10.xml`` rather than after it as it does in a
    lexicographical sort.
    """
    def __init__(self):
        super(PartCollection, self).__init__()

    def _loadpart(self, part):
        """
        Insert a new part loaded from a package, such that list remains
        sorted in logical partname order (e.g. slide10.xml comes after
        slide9.xml).
        """
        new_partidx = util.Partname(part.partname).idx
        for idx, seq_part in enumerate(self._values):
            partidx = util.Partname(seq_part.partname).idx
            if partidx > new_partidx:
                self._values.insert(idx, part)
                return
        self._values.append(part)


class ImageCollection(PartCollection):
    """
    Immutable sequence of images, typically belonging to an instance of
    :class:`Package`. An image part containing a particular image blob appears
    only once in an instance, regardless of how many times it is referenced by
    a pic shape in a slide.
    """
    def __init__(self):
        super(ImageCollection, self).__init__()

    def add_image(self, file):
        """
        Return image part containing the image in *file*, which is either a
        path to an image file or a file-like object containing an image. If an
        image instance containing this same image already exists, that
        instance is returned. If it does not yet exist, a new one is created.
        """
        # use Image constructor to validate and characterize image file
        image = Image(file)
        # return matching image if found
        for existing_image in self._values:
            if existing_image._sha1 == image._sha1:
                return existing_image
        # otherwise add it to collection and return new image
        self._values.append(image)
        self.__rename_images()
        return image

    def __rename_images(self):
        """
        Assign partnames like ``/ppt/media/image9.png`` to all images in the
        collection. The name portion is always ``image``. The number part
        forms a continuous sequence starting at 1 (e.g. 1, 2, 3, ...). The
        extension is preserved during renaming.
        """
        for idx, image in enumerate(self._values):
            image.partname = '/ppt/media/image%d%s' % (idx+1, image.ext)


class Part(object):
    """
    Part factory. Returns an instance of the appropriate custom part type for
    part types that have them, BasePart otherwise.
    """
    def __new__(cls, reltype, content_type):
        """
        *reltype* is the relationship type, e.g. ``RT_SLIDE``, corresponding
        to the type of part to be created. For at least one part type, in
        particular the presentation part type, *content_type* is also required
        in order to fully specify the part to be created.
        """
        # log.debug("Creating Part for %s", reltype)
        if reltype == RT_OFFICE_DOCUMENT:
            if content_type in (CT_PRESENTATION, CT_TEMPLATE, CT_SLIDESHOW):
                return Presentation()
            else:
                tmpl = "Not a presentation content type, got '%s'"
                raise InvalidPackageError(tmpl % content_type)
        elif reltype == RT_SLIDE:
            return Slide()
        elif reltype == RT_SLIDE_LAYOUT:
            return SlideLayout()
        elif reltype == RT_SLIDE_MASTER:
            return SlideMaster()
        elif reltype == RT_IMAGE:
            return Image()
        return BasePart()


class BasePart(Observable):
    """
    Base class for presentation model parts. Provides common code to all parts
    and is the class we instantiate for parts we don't unmarshal or manipulate
    yet.

    .. attribute:: _element

       ElementTree element for XML parts. ``None`` for binary parts.

    .. attribute:: _load_blob

       Contents of part as a byte string extracted from the package file. May
       be set to ``None`` by subclasses that override ._blob after content is
       unmarshaled, to free up memory.

    .. attribute:: _relationships

       :class:`RelationshipCollection` instance containing the relationships
       for this part.

    """
    def __init__(self, content_type=None):
        """
        Needs content_type parameter so newly created parts (not loaded from
        package) can register their content type.
        """
        super(BasePart, self).__init__()
        self.__content_type = content_type
        self.__partname = None
        self._element = None
        self._load_blob = None
        self._relationships = _RelationshipCollection()

    @property
    def _blob(self):
        """
        Default is to return unchanged _load_blob. Dynamic parts will
        override. Raises :class:`ValueError` if _load_blob is None.
        """
        if self.partname.endswith('.xml'):
            assert self._element is not None, 'BasePart._blob is undefined '\
                'for xml parts when part.__element is None'
            xml = oxml_tostring(self._element, encoding='UTF-8',
                                pretty_print=True, standalone=True)
            return xml
        # default for binary parts is to return _load_blob unchanged
        assert self._load_blob, "BasePart._blob called on part with no "\
            "_load_blob; perhaps _blob not overridden by sub-class?"
        return self._load_blob

    @property
    def _content_type(self):
        """
        Content type of this part, e.g.
        'application/vnd.openxmlformats-officedocument.theme+xml'.
        """
        if self.__content_type is None:
            msg = "BasePart._content_type accessed before assigned"
            raise ValueError(msg)
        return self.__content_type

    @_content_type.setter
    def _content_type(self, content_type):
        self.__content_type = content_type

    @property
    def partname(self):
        """Part name of this part, e.g. '/ppt/slides/slide1.xml'."""
        assert self.__partname, "BasePart.partname referenced before assigned"
        return self.__partname

    @partname.setter
    def partname(self, partname):
        self.__partname = partname
        self._notify_observers('partname', self.__partname)

    def _add_relationship(self, reltype, target_part):
        """
        Return new relationship of *reltype* to *target_part* after adding it
        to the relationship collection of this part.
        """
        rId = self._relationships._next_rId
        rel = _Relationship(rId, reltype, target_part)
        self._relationships._additem(rel)
        return rel

    def _load(self, pkgpart, part_dict):
        """
        Load part and relationships from package part, and propagate load
        process down the relationship graph. *pkgpart* is an instance of
        :class:`pptx.packaging.Part` containing the part contents read from
        the on-disk package. *part_dict* is a dictionary of already-loaded
        parts, keyed by partname.
        """
        # log.debug("loading part %s", pkgpart.partname)

        # # set attributes from package part
        self.__content_type = pkgpart.content_type
        self.__partname = pkgpart.partname
        if pkgpart.partname.endswith('.xml'):
            self._element = oxml_fromstring(pkgpart.blob)
        else:
            self._load_blob = pkgpart.blob

        # discard any previously loaded relationships
        self._relationships = _RelationshipCollection()

        # load relationships and propagate load for related parts
        for pkgrel in pkgpart.relationships:
            # unpack working values for part to be loaded
            reltype = pkgrel.reltype
            target_pkgpart = pkgrel.target
            partname = target_pkgpart.partname
            content_type = target_pkgpart.content_type

            # create target part
            if partname in part_dict:
                part = part_dict[partname]
            else:
                part = Part(reltype, content_type)
                part_dict[partname] = part
                part._load(target_pkgpart, part_dict)

            # create model-side package relationship
            model_rel = _Relationship(pkgrel.rId, reltype, part)
            self._relationships._additem(model_rel)
        return self


class Presentation(BasePart):
    """
    Top level class in object model, represents the contents of the /ppt
    directory of a .pptx file.
    """
    def __init__(self):
        super(Presentation, self).__init__()
        self.__slidemasters = PartCollection()
        self.__slides = SlideCollection(self)

    @property
    def slidemasters(self):
        """
        List of :class:`SlideMaster` objects belonging to this presentation.
        """
        return tuple(self.__slidemasters)

    @property
    def slides(self):
        """
        :class:`SlideCollection` object containing the slides in this
        presentation.
        """
        return self.__slides

    @property
    def _blob(self):
        """
        Rewrite sldId elements in sldIdLst before handing over to super for
        transformation of _element into a blob.
        """
        self.__rewrite_sldIdLst()
        # # at least the following needs to be added before using
        # # _reltype_ordering again for Presentation
        # self.__rewrite_notesMasterIdLst()
        # self.__rewrite_handoutMasterIdLst()
        # self.__rewrite_sldMasterIdLst()
        return super(Presentation, self)._blob

    def _load(self, pkgpart, part_dict):
        """
        Load presentation from package part.
        """
        # call parent to do generic aspects of load
        super(Presentation, self)._load(pkgpart, part_dict)

        # side effect of setting reltype ordering is that rId values can be
        # changed (renumbered during resequencing), so must complete rewrites
        # of all four IdLst elements (notesMasterIdLst, etc.) internal to
        # presentation.xml to reflect any possible changes. Not sure if good
        # order in the .rels files is worth the trouble just yet, so
        # commenting this out for now.

        # # set reltype ordering so rels file ordering is readable
        # self._relationships._reltype_ordering = (RT_SLIDE_MASTER,
        #     RT_NOTES_MASTER, RT_HANDOUT_MASTER, RT_SLIDE,
        #     RT_PRES_PROPS, RT_VIEW_PROPS, RT_TABLE_STYLES, RT_THEME)

        # selectively unmarshal relationships for now
        for rel in self._relationships:
            # log.debug("Presentation Relationship %s", rel._reltype)
            if rel._reltype == RT_SLIDE_MASTER:
                self.__slidemasters._loadpart(rel._target)
            elif rel._reltype == RT_SLIDE:
                self.__slides._loadpart(rel._target)
        return self

    def __rewrite_sldIdLst(self):
        """
        Rewrite the ``<p:sldIdLst>`` element in ``<p:presentation>`` to
        reflect current ordering of slide relationships and possible
        renumbering of ``rId`` values.
        """
        sldIdLst = _child(self._element, 'p:sldIdLst', _nsmap)
        if sldIdLst is None:
            sldIdLst = self.__add_sldIdLst()
        sldIdLst.clear()
        sld_rels = self._relationships.rels_of_reltype(RT_SLIDE)
        for idx, rel in enumerate(sld_rels):
            sldId = _Element('p:sldId', _nsmap)
            sldIdLst.append(sldId)
            sldId.set('id', str(256+idx))
            sldId.set(qn('r:id'), rel._rId)

    def __add_sldIdLst(self):
        """
        Add a <p:sldIdLst> element to <p:presentation> in the right sequence
        among its siblings.
        """
        sldIdLst = _child(self._element, 'p:sldIdLst', _nsmap)
        assert sldIdLst is None, '__add_sldIdLst() called where '\
                                 '<p:sldIdLst> already exists'
        sldIdLst = _Element('p:sldIdLst', _nsmap)
        # insert new sldIdLst element in right sequence
        sldSz = _child(self._element, 'p:sldSz', _nsmap)
        if sldSz is not None:
            sldSz.addprevious(sldIdLst)
        else:
            notesSz = _child(self._element, 'p:notesSz', _nsmap)
            notesSz.addprevious(sldIdLst)
        return sldIdLst


class Image(BasePart):
    """
    Return new Image part instance. *file* may be |None|, a path to a file (a
    string), or a file-like object. If *file* is |None|, no image is loaded
    and :meth:`_load` must be called before using the instance. Otherwise, the
    file referenced or contained in *file* is loaded. Corresponds to package
    files ppt/media/image[1-9][0-9]*.*.
    """
    def __init__(self, file=None):
        super(Image, self).__init__()
        self.__ext = None
        if file is not None:
            self.__load_image_from_file(file)

    @property
    def ext(self):
        """Return file extension for this image"""
        assert self.__ext, "Image.__ext referenced before assigned"
        return self.__ext

    @property
    def _sha1(self):
        """Return SHA1 hash digest for image"""
        return hashlib.sha1(self._blob).hexdigest()

    @property
    def _blob(self):
        """
        For an image, _blob is always _load_blob, image file content is not
        manipulated.
        """
        return self._load_blob

    def _load(self, pkgpart, part_dict):
        """Handle aspects of loading that are particular to image parts."""
        # call parent to do generic aspects of load
        super(Image, self)._load(pkgpart, part_dict)
        # set file extension
        self.__ext = posixpath.splitext(pkgpart.partname)[1]
        # return self-reference to allow generative calling
        return self

    @staticmethod
    def __image_ext_content_type(ext):
        """Return the content type corresponding to filename extension *ext*"""
        if ext not in spec.default_content_types:
            tmpl = "unsupported image file extension '%s'"
            raise TypeError(tmpl % (ext))
        content_type = spec.default_content_types[ext]
        if not content_type.startswith('image/'):
            tmpl = "'%s' is not an image content type; ext '%s'"
            raise TypeError(tmpl % (content_type, ext))
        return content_type

    @staticmethod
    def __ext_from_image_stream(stream):
        """
        Return the filename extension appropriate to the image file contained
        in *stream*.
        """
        ext_map = {'GIF': '.gif', 'JPEG': '.jpg', 'PNG': '.png',
                   'TIFF': '.tiff', 'WMF': '.wmf'}
        stream.seek(0)
        format = PIL_Image.open(stream).format
        if format not in ext_map:
            tmpl = "unsupported image format, expected one of: %s, got '%s'"
            raise ValueError(tmpl % (ext_map.keys(), format))
        return ext_map[format]

    def __load_image_from_file(self, file):
        """
        Load image from *file*, which is either a path to an image file or a
        file-like object.
        """
        if isinstance(file, basestring):  # file is a path
            path = file
            self.__ext = os.path.splitext(path)[1]
            self._content_type = self.__image_ext_content_type(self.__ext)
            with open(path, 'rb') as f:
                self._load_blob = f.read()
        else:  # assume file is a file-like object
            self.__ext = self.__ext_from_image_stream(file)
            self._content_type = self.__image_ext_content_type(self.__ext)
            file.seek(0)
            self._load_blob = file.read()


# ============================================================================
# Slide Parts
# ============================================================================

class SlideCollection(PartCollection):
    """
    Immutable sequence of slides belonging to an instance of |Presentation|,
    with methods for manipulating the slides in the presentation.
    """
    def __init__(self, presentation):
        super(SlideCollection, self).__init__()
        self.__presentation = presentation

    def add_slide(self, slidelayout):
        """Add a new slide that inherits layout from *slidelayout*."""
        # 1. construct new slide
        slide = Slide(slidelayout)
        # 2. add it to this collection
        self._values.append(slide)
        # 3. assign its partname
        self.__rename_slides()
        # 4. add presentation->slide relationship
        self.__presentation._add_relationship(RT_SLIDE, slide)
        # 5. return reference to new slide
        return slide

    def __rename_slides(self):
        """
        Assign partnames like ``/ppt/slides/slide9.xml`` to all slides in the
        collection. The name portion is always ``slide``. The number part
        forms a continuous sequence starting at 1 (e.g. 1, 2, 3, ...). The
        extension is always ``.xml``.
        """
        for idx, slide in enumerate(self._values):
            slide.partname = '/ppt/slides/slide%d.xml' % (idx+1)


class BaseSlide(BasePart):
    """
    Base class for slide parts, e.g. slide, slideLayout, slideMaster,
    notesSlide, notesMaster, and handoutMaster.
    """
    def __init__(self, content_type=None):
        """
        Needs content_type parameter so newly created parts (not loaded from
        package) can register their content type.
        """
        super(BaseSlide, self).__init__(content_type)
        self._shapes = None

    @property
    def name(self):
        """Internal name of this slide-like object."""
        cSld = self._element.cSld
        return cSld.get('name', default='')

    @property
    def shapes(self):
        """Collection of shape objects belonging to this slide."""
        assert self._shapes is not None, ("BaseSlide.shapes referenced "
                                          "before assigned")
        return self._shapes

    def _load(self, pkgpart, part_dict):
        """Handle aspects of loading that are general to slide types."""
        # call parent to do generic aspects of load
        super(BaseSlide, self)._load(pkgpart, part_dict)
        # unmarshal shapes
        self._shapes = ShapeCollection(self._element.cSld.spTree, self)
        # return self-reference to allow generative calling
        return self

    @property
    def _package(self):
        """Reference to |Package| containing this slide"""
        return Package.containing(self)


class Slide(BaseSlide):
    """
    Slide part. Corresponds to package files ppt/slides/slide[1-9][0-9]*.xml.
    """
    def __init__(self, slidelayout=None):
        super(Slide, self).__init__(CT_SLIDE)
        self.__slidelayout = slidelayout
        self._element = self.__minimal_element
        self._shapes = ShapeCollection(self._element.cSld.spTree, self)
        # if slidelayout, this is a slide being added, not one being loaded
        if slidelayout:
            self._shapes._clone_layout_placeholders(slidelayout)
            # add relationship to slideLayout part
            self._add_relationship(RT_SLIDE_LAYOUT, slidelayout)

    @property
    def slidelayout(self):
        """
        :class:`SlideLayout` object this slide inherits appearance from.
        """
        return self.__slidelayout

    def _load(self, pkgpart, part_dict):
        """
        Load slide from package part.
        """
        # call parent to do generic aspects of load
        super(Slide, self)._load(pkgpart, part_dict)
        # selectively unmarshal relationships for now
        for rel in self._relationships:
            # log.debug("SlideMaster Relationship %s", rel._reltype)
            if rel._reltype == RT_SLIDE_LAYOUT:
                self.__slidelayout = rel._target
        return self

    @property
    def __minimal_element(self):
        """
        Return element containing the minimal XML for a slide, based on what
        is required by the XMLSchema.
        """
        sld = _Element('p:sld', _nsmap)
        _SubElement(sld, 'p:cSld', _nsmap)
        _SubElement(sld.cSld, 'p:spTree', _nsmap)
        _SubElement(sld.cSld.spTree, 'p:nvGrpSpPr', _nsmap)
        _SubElement(sld.cSld.spTree.nvGrpSpPr, 'p:cNvPr', _nsmap)
        _SubElement(sld.cSld.spTree.nvGrpSpPr, 'p:cNvGrpSpPr', _nsmap)
        _SubElement(sld.cSld.spTree.nvGrpSpPr, 'p:nvPr', _nsmap)
        _SubElement(sld.cSld.spTree, 'p:grpSpPr', _nsmap)
        sld.cSld.spTree.nvGrpSpPr.cNvPr.set('id', '1')
        sld.cSld.spTree.nvGrpSpPr.cNvPr.set('name', '')
        return sld


class SlideLayout(BaseSlide):
    """
    Slide layout part. Corresponds to package files
    ppt/slideLayouts/slideLayout[1-9][0-9]*.xml.
    """
    def __init__(self):
        super(SlideLayout, self).__init__(CT_SLIDE_LAYOUT)
        self.__slidemaster = None

    @property
    def slidemaster(self):
        """Slide master from which this slide layout inherits properties."""
        assert self.__slidemaster is not None, ("SlideLayout.slidemaster "
                                                "referenced before assigned")
        return self.__slidemaster

    def _load(self, pkgpart, part_dict):
        """
        Load slide layout from package part.
        """
        # call parent to do generic aspects of load
        super(SlideLayout, self)._load(pkgpart, part_dict)

        # selectively unmarshal relationships we need
        for rel in self._relationships:
            # log.debug("SlideLayout Relationship %s", rel._reltype)
            # get slideMaster from which this slideLayout inherits properties
            if rel._reltype == RT_SLIDE_MASTER:
                self.__slidemaster = rel._target

        # return self-reference to allow generative calling
        return self


class SlideMaster(BaseSlide):
    """
    Slide master part. Corresponds to package files
    ppt/slideMasters/slideMaster[1-9][0-9]*.xml.
    """
    # TECHNOTE: In the Microsoft API, Master is a general type that all of
    # SlideMaster, SlideLayout (CustomLayout), HandoutMaster, and NotesMaster
    # inherit from. So might look into why that is and consider refactoring
    # the various masters a bit later.

    def __init__(self):
        super(SlideMaster, self).__init__(CT_SLIDE_MASTER)
        self.__slidelayouts = PartCollection()

    @property
    def slidelayouts(self):
        """
        Collection of slide layout objects belonging to this slide master.
        """
        return self.__slidelayouts

    def _load(self, pkgpart, part_dict):
        """
        Load slide master from package part.
        """
        # call parent to do generic aspects of load
        super(SlideMaster, self)._load(pkgpart, part_dict)

        # selectively unmarshal relationships for now
        for rel in self._relationships:
            # log.debug("SlideMaster Relationship %s", rel._reltype)
            if rel._reltype == RT_SLIDE_LAYOUT:
                self.__slidelayouts._loadpart(rel._target)
        return self
