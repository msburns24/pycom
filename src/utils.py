from __future__ import annotations
from typing import Any, Callable


def extract_attributes(
        to_object: object,
        from_object: object,
        attrs_map: dict[str, str],
) -> None:
    '''
    Extracts attributes specified in `attrs_map` and adds them to the provided
    object. When an exception is raised, stores `None`.

    Parameters
    ----------
    to_object : object
        The object in which to store the extracted attributes.
    from_object : object
        The object from which to extract the attributes
    attrs_map : dict[str, str]
        A map of `from_attr_name` to `to_attr_name`
    
    Returns
    -------
    None
    '''
    for from_attr_name, to_attr_name in attrs_map.items():
        try:
            value = getattr(from_object, from_attr_name)
        except:
            value = None
        setattr(to_object, to_attr_name, value)
    return