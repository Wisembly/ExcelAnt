<?php

namespace ExcelAnt\Collections;

use ArrayIterator;

use ExcelAnt\Collections\StyleCollectionInterface;
use ExcelAnt\Style\StyleInterface;

class StyleCollection implements StyleCollectionInterface
{
    private $_styles = [];

    /**
     * @param array $styles Array of StyleInterface
     */
    public function __construct(array $styles)
    {
        foreach ($styles as $style) {
            $this->add($style);
        }
    }

    /**
     * {@inheritDoc}
     */
    public function add(StyleInterface $style)
    {
        if (false !== $key = $this->contains($style)) {
            $this->set($key, $style);

            return;
        }

        $this->_styles[] = $style;
    }

    /**
     * {@inheritDoc}
     */
    public function set($key, StyleInterface $style)
    {
        $keyExist = $this->contains($style);

        if (false !== $keyExist && ($keyExist !== $key)) {
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
    public function removeElement(StyleInterface $style)
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
        if (!isset($this->_styles[$key])) {
            throw new \OutOfBoundsException("Index doesn't exist");
        }

        return $this->_styles[$key];
    }

    /**
     * {@inheritDoc}
     */
    public function getElement(StyleInterface $style)
    {
        $key = $this->contains($style);

        if (false === $key) {
            throw new \OutOfBoundsException("This style doesn't exist in the styleCollection");
        }

        return $this->get($key);
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
        return $this->add($value);
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
        return new ArrayIterator($this->_styles);
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
}