<?php

namespace avadim\FastExcelWriter\Charts;

class DataSeriesSet
{
    protected $dataValues;

    protected $dataLabels;

    protected $dataCategories;

    protected $dataOptions;


    public function __construct($dataValues, $dataLabels, $dataCategories, $dataOptions)
    {
        $this->dataValues = $dataValues;
        $this->dataLabels = $dataLabels;
        $this->dataCategories = $dataCategories;
        $this->dataOptions = $dataOptions;
    }


    public function getDataValues()
    {
        return $this->dataValues;
    }


    public function getDataLabels()
    {
        return $this->dataLabels;
    }


    public function getDataCategories()
    {
        return $this->dataCategories;
    }


    public function getDataOptions()
    {
        return $this->dataOptions;
    }
}