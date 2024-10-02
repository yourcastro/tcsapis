import React, { useState, useMemo } from "react";
import { Container, Row, Col, Table, Pagination } from "react-bootstrap";

const GridWithOptimizedPagination = () => {
  // Example data
  const data = useMemo(() => Array.from({ length: 100 }, (_, index) => `Item ${index + 1}`), []);
  
  // Pagination state
  const [currentPage, setCurrentPage] = useState(1);
  const itemsPerPage = 10;
  const totalPages = Math.ceil(data.length / itemsPerPage);
  const paginationRange = 3; // Number of pagination links to display around the current page
  
  // Get current items for pagination
  const currentData = useMemo(() => {
    const start = (currentPage - 1) * itemsPerPage;
    return data.slice(start, start + itemsPerPage);
  }, [currentPage, data]);

  // Pagination change handler
  const handlePageChange = (pageNumber) => {
    setCurrentPage(pageNumber);
  };

  // Helper function to create pagination items
  const generatePaginationItems = () => {
    const items = [];

    // Always show the first page link
    if (currentPage > 1) {
      items.push(
        <Pagination.First
          key="first"
          onClick={() => handlePageChange(1)}
        />
      );
      items.push(
        <Pagination.Prev
          key="prev"
          onClick={() => handlePageChange(currentPage - 1)}
          disabled={currentPage === 1}
        />
      );
    }

    // Generate page numbers with dynamic range
    const startPage = Math.max(1, currentPage - paginationRange);
    const endPage = Math.min(totalPages, currentPage + paginationRange);

    for (let i = startPage; i <= endPage; i++) {
      items.push(
        <Pagination.Item
          key={i}
          active={i === currentPage}
          onClick={() => handlePageChange(i)}
        >
          {i}
        </Pagination.Item>
      );
    }

    // Always show the last page link
    if (currentPage < totalPages) {
      items.push(
        <Pagination.Next
          key="next"
          onClick={() => handlePageChange(currentPage + 1)}
          disabled={currentPage === totalPages}
        />
      );
      items.push(
        <Pagination.Last
          key="last"
          onClick={() => handlePageChange(totalPages)}
        />
      );
    }

    return items;
  };

  return (
    <Container>
      {/* Grid Header */}
      <Row className="my-3">
        <Col>
          <h3>Grid Header</h3>
        </Col>
      </Row>

      {/* Grid Content */}
      <Row>
        <Col>
          <Table striped bordered hover>
            <thead>
              <tr>
                <th>#</th>
                <th>Item</th>
              </tr>
            </thead>
            <tbody>
              {currentData.map((item, index) => (
                <tr key={index}>
                  <td>{(currentPage - 1) * itemsPerPage + index + 1}</td>
                  <td>{item}</td>
                </tr>
              ))}
            </tbody>
          </Table>
        </Col>
      </Row>

      {/* Optimized Pagination */}
      <Row>
        <Col>
          <Pagination>{generatePaginationItems()}</Pagination>
        </Col>
      </Row>
    </Container>
  );
};

export default GridWithOptimizedPagination;
