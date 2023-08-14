
/**
 * Count objects by number of instances of unique values of key
 * @param {Object[]} objects 
 * @param {string} key 
 */
function countBy(objects, key)
{
  return objects.reduce(function (counts, object)
  {
    if (object[key] !== undefined)
    {
      counts[object[key]] = counts[object[key]] || 0
      counts[object[key]] += 1
      counts.total += 1;
    }
    return counts;
  }, { 'total': 0 })
}