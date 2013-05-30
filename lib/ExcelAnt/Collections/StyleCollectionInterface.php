<?php

namespace ExcelAnt\Collections;

use Closure, Countable, IteratorAggregate, ArrayAccess;

use ExcelAnt\Style\StyleInterface;

interface StyleCollectionInterface extends Countable, IteratorAggregate, ArrayAccess
{
    /**
     * Adds a style at the end of the collection.
     * If the style already exist in the colleciton, replace it.
     *
     * @param StyleInterface $style The style to add.
     */
    public function add(StyleInterface $style);

    /**
     * Sets a style in the collection at the specified key/index if there isn't another identic style.
     * Else throw InvalidException
     *
     * @param string|integer $key   The key/index of the style to set.
     * @param mixed          $value The style to set.
     *
     * @throws InvalidException If there is another identic style in the styleCollection
     */
    public function set($key, StyleInterface $value);

    /**
     * Checks whether a style is contained in the collection.
     *
     * @param StyleInterface $style The style to search for.
     *
     * @return mixed Key if the collection contains the style, FALSE otherwise.
     */
    public function contains(StyleInterface $style);

    /**
     * Clears the collection, removing all styles.
     */
    public function clear();

    /**
     * Checks whether the collection is empty (contains no styles).
     *
     * @return boolean TRUE if the collection is empty, FALSE otherwise.
     */
    public function isEmpty();

    /**
     * Removes the style at the specified index from the collection.
     *
     * @param string|integer $key The kex/index of the style to remove.
     *
     * @return mixed The removed style or NULL, if the collection did not contain the style.
     */
    public function remove($key);

    /**
     * Removes the specified style from the collection, if it is found.
     *
     * @param StyleInterface $style The style to remove.
     *
     * @return boolean TRUE if this collection contained the specified style, FALSE otherwise.
     */
    public function removeElement(StyleInterface $style);

    /**
     * Checks whether the collection contains a style with the specified key/index.
     *
     * @param string|integer $key The key/index to check for.
     *
     * @return boolean TRUE if the collection contains a style with the specified key/index,
     *                 FALSE otherwise.
     */
    public function containsKey($key);

    /**
     * Gets the style at the specified key/index.
     *
     * @param string|integer $key The key/index of the style to retrieve.
     *
     * @throws OutOfBoundsException If there is no index with key value
     *
     * @return mixed
     */
    public function get($key);

    /**
     * Gets all keys/indices of the collection.
     *
     * @return array The keys/indices of the collection, in the order of the corresponding
     *               styles in the collection.
     */
    public function getKeys();

    /**
     * Gets a native PHP array representation of the collection.
     *
     * @return array
     */
    public function toArray();
}