<?php

namespace ExcelAnt\Collections;

use ArrayIterator;
use ExcelAnt\Style\StyleInterface;

class StyleCollection implements Collection
{
    private $_styles;

    public function __construct($styles)
    {
        if (isset($styles)) {
           if (!is_array($styles)) {
               throw new \InvalidArgumentException("styles must be an array of StyleInterface");
           }
           foreach ($styles as $style) {
               if (!($style instanceof StyleInterface)) {
                   throw new \InvalidArgumentException("style must be an instance of StyleInterface");
               }
           }
       }

        $this->_styles = $styles;
    }

    /**
     * {@inheritDoc}
     */
    public function add(StyleInterface $style)
    {
        if (false !== $key = $this->contains($style)) {
            $this->hardSet($key, $style);

            return;
        }

        $this->_styles[] = $style;
    }

    /**
     * {@inheritDoc}
     */
    public function set($key, StyleInterface $style)
    {
        if (false !== $key = $this->contains($style)) {
            throw new \InvalidArgumentException("This style already exist in the styleCollection");
        }

        $this->_styles[$key] = $style;
    }

    /**
     * {@inheritDoc}
     */
    public function contains(StyleInterface $style)
    {
        foreach ($this->_styles as $key => $elem) {
            if ($elem instanceof $style) {
                return $key;
            }
        }

        return false;
    }

    /**
     * {@inheritDoc}
     */
    public function clear()
    {
        $this->_styles = array();
    }

    /**
     * {@inheritDoc}
     */
    public function isEmpty()
    {
        return !$this->_styles;
    }

    /**
     * {@inheritDoc}
     */
    public function remove($key)
    {
        if (isset($this->_styles[$key]) || array_key_exists($key, $this->_styles)) {
            $removed = $this->_styles[$key];
            unset($this->_styles[$key]);

            return $removed;
        }

        return null;
    }

    /**
     * {@inheritDoc}
     */
    public function removeElement($style)
    {
        $key = $this->contains($style);

        if ($key !== false) {
            unset($this->_styles[$key]);

            return true;
        }

        return false;
    }

    /**
     * {@inheritDoc}
     */
    public function containsKey($key)
    {
        return isset($this->_styles[$key]) || array_key_exists($key, $this->_styles);
    }

    /**
     * {@inheritDoc}
     */
    public function get($key)
    {
        if (isset($this->_styles[$key])) {
            return $this->_styles[$key];
        }

        return null;
    }

    /**
     * {@inheritDoc}
     */
    public function getKeys()
    {
        return array_keys($this->_styles);
    }

    /**
     * {@inheritDoc}
     */
    public function toArray()
    {
        return $this->_styles;
    }

    /**
     * {@inheritDoc}
     */
    public function first()
    {
        return reset($this->_styles);
    }

    /**
     * {@inheritDoc}
     */
    public function last()
    {
        return end($this->_styles);
    }

    /*
     * Required by interface ArrayAccess.
     *
     * {@inheritDoc}
     */
    public function offsetExists($offset)
    {
        return $this->containsKey($offset);
    }

    /*
     * Required by interface ArrayAccess.
     *
     * {@inheritDoc}
     */
    public function offsetGet($offset)
    {
        return $this->get($offset);
    }

    /*
     * Required by interface ArrayAccess.
     *
     * {@inheritDoc}
     */
    public function offsetSet($offset, $value)
    {
        if (!isset($offset)) {
            return $this->add($value);
        }

        return $this->set($offset, $value);
    }

    /*
     * Required by interface ArrayAccess.
     *
     * {@inheritDoc}
     */
    public function offsetUnset($offset)
    {
        return $this->remove($offset);
    }

    /*
     * Required by interface Iterator.
     *
     * {@inheritDoc}
     */
    public function getIterator()
    {
        return new ArrayIterator($this->styles);
    }

    /*
     * Required by interface Countable.
     *
     * {@inheritDoc}
     */
    public function count()
    {
        return count($this->_styles);
    }

    /**
     * Used to replace a style in styleCollection
     *
     * @param  int            $key   The index
     * @param  StyleInterface $style The style
     */
    private function hardSet($key, StyleInterface $style)
    {
        $this->_styles[$key] = $style;
    }
}