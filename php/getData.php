<?php

require 'vendor/autoload.php';

use Microsoft\Graph\Graph;
use Microsoft\Graph\Model;

class Events
{
    public function getEvents($startDate = '2018-12-01', $endDate = '2018-12-04')
    {
        // NEEDS TO BE POPULATED
        $accessToken = '';

        $graph = new Graph();
        $graph->setAccessToken($accessToken);

        $request = "/me/calendar/events?filter=start/dateTime ge '".$startDate."T00:00' and end/dateTime lt '".$endDate."T00:00'&select=subject,categories,start,end&top=500";
        $events = $graph->createRequest("GET", $request)
                      ->setReturnType(Model\Event::class)
                      ->execute();

        return $events;
    }
}

$events = new Events();

$weeks = [];
for ($i=4;$i>0;$i--) {
    $weeks['-' . $i] = [
        'start' => date('Y-m-d', strtotime('-'.$i.' Monday')),
        'end' => date('Y-m-d', strtotime('-'.$i.' Saturday'))
    ];
}

$categoryList = [];
$data = [];
foreach ($weeks as $key => $value) {
    $data[$key] = [];
    $currentEvents = $events->getEvents($value['start'], $value['end']);
    foreach ($currentEvents as $currentEvent) {
        $start = $currentEvent->getStart()->getDateTime();
        $end = $currentEvent->getEnd()->getDateTime();
        $duration = ((strtotime($end) - strtotime($start)) / 60) / 60; // hours;
        $categories = $currentEvent->getCategories();
        $category = "None";
        if (!empty($categories)) {
            $category = $categories[0];
        }
        if (isset($data[$key][$category])) {
            $data[$key][$category] += $duration;
        }
        if (!isset($data[$key][$category])) {
            $data[$key][$category] = $duration;
        }
        if (!in_array($category, $categoryList)) {
            $categoryList[] = $category;
        }
    };
}

sort($categoryList);

foreach ($data as $week => $categories) {
    foreach ($categoryList as $category) {
        if (!array_key_exists($category, $categories)) {
            $categories[$category] = 0;
        }
    }
    ksort($categories);
    $data[$week] = $categories;
}

$result = [];

$result['cols'] = [];
$result['cols'][] = [
    'label' => 'weeks', 
    'type' => 'number'
];

foreach ($categoryList as $category) {
    $result['cols'][] = [
        'label' => $category, 
        'type' => 'number'
    ];
}

$rows = [];
foreach ($data as $week => $details) {
    $sub_array = [];
    $sub_array[] =  array(
        "v" => $week
    );
    foreach ($details as $hours) {
        $sub_array[] =  array(
            "v" => $hours
        );
    }
    $rows[] =  [
        "c" => $sub_array
    ];
}
$result["rows"] = $rows;

echo json_encode($result);
