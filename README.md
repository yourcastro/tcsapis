import React, { useEffect } from 'react';
import { useDispatch } from 'react-redux';
import { fetchData } from './actions'; // Assume this is your action creator

const FetchDataComponent = () => {
  const dispatch = useDispatch();

  useEffect(() => {
    const fetchDataInLoop = async () => {
      const ids = [1, 2, 3, 4, 5]; // Example IDs to fetch

      try {
        // Dispatch API calls in parallel
        await Promise.all(ids.map(id => dispatch(fetchData(id))));
      } catch (error) {
        console.error('Error fetching data:', error);
      }
    };

    fetchDataInLoop();
  }, [dispatch]);

  return <div>Data is being fetched...</div>;
};

export default FetchDataComponent;
