<?php
if(!function_exists('is_assoc')) {
    function is_assoc(array $arr) {
        if (array() === $arr) return false;
        return array_keys($arr) !== range(0, count($arr) - 1);
    }
}