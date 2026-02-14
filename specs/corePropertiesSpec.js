const { describe, it, beforeEach } = require('node:test');
const assert = require('node:assert/strict');
var xl = require('../lib/libxl'),
    util = require('util'),
    testUtils = require('./testUtils'),
    shouldThrow = testUtils.shouldThrow;

describe('The CoreProperties class', () => {
    let book;
    let coreProperties;

    beforeEach(() => {
        book = new xl.Book(xl.BOOK_TYPE_XLSX);
        coreProperties = book.coreProperties();
    });

    it('title and setTitle change the title', () => {
        shouldThrow(coreProperties.title, {});
        shouldThrow(coreProperties.setTitle, coreProperties, 1);
        shouldThrow(coreProperties.setTitle, {}, 'Foo');

        assert.strictEqual(coreProperties.setTitle('Foo'), coreProperties);
        assert.strictEqual(coreProperties.title(), 'Foo');
    });

    it('subject and setSubject change the title', () => {
        shouldThrow(coreProperties.subject, {});
        shouldThrow(coreProperties.setSubject, coreProperties, 1);
        shouldThrow(coreProperties.setSubject, {}, 'Foo');

        assert.strictEqual(coreProperties.setSubject('Foo'), coreProperties);
        assert.strictEqual(coreProperties.subject(), 'Foo');
    });

    it('creator and setCreator change the title', () => {
        shouldThrow(coreProperties.creator, {});
        shouldThrow(coreProperties.setCreator, coreProperties, 1);
        shouldThrow(coreProperties.setCreator, {}, 'Foo');

        assert.strictEqual(coreProperties.setCreator('Foo'), coreProperties);
        assert.strictEqual(coreProperties.creator(), 'Foo');
    });

    it('lastModifiedBy and setLastModifiedBy change the title', () => {
        shouldThrow(coreProperties.lastModifiedBy, {});
        shouldThrow(coreProperties.setLastModifiedBy, coreProperties, 1);
        shouldThrow(coreProperties.setLastModifiedBy, {}, 'Foo');

        assert.strictEqual(coreProperties.setLastModifiedBy('Foo'), coreProperties);
        assert.strictEqual(coreProperties.lastModifiedBy(), 'Foo');
    });

    it('created and setCreated change the title', () => {
        const now = new Date().toISOString();

        shouldThrow(coreProperties.created, {});
        shouldThrow(coreProperties.setCreated, coreProperties, 1);
        shouldThrow(coreProperties.setCreated, {}, now);

        let date = new Date();
        assert.strictEqual(coreProperties.setCreated(now), coreProperties);
        assert.strictEqual(coreProperties.created(), now);
    });

    it('createdAsDouble and setCreatedAsDouble change the title', () => {
        shouldThrow(coreProperties.createdAsDouble, {});
        shouldThrow(coreProperties.setCreatedAsDouble, coreProperties, 'a');
        shouldThrow(coreProperties.setCreatedAsDouble, {}, 1.0);

        assert.strictEqual(coreProperties.setCreatedAsDouble(1.0), coreProperties);
        assert.ok(coreProperties.createdAsDouble() - 1.0 < testUtils.epsilon);
    });

    it('modified and setModified change the title', () => {
        const now = new Date().toISOString();

        shouldThrow(coreProperties.modified, {});
        shouldThrow(coreProperties.setModified, coreProperties, 1);
        shouldThrow(coreProperties.setModified, {}, now);

        let date = new Date();
        assert.strictEqual(coreProperties.setModified(now), coreProperties);
        assert.strictEqual(coreProperties.modified(), now);
    });

    it('modifiedAsDouble and setModifiedAsDouble change the title', () => {
        shouldThrow(coreProperties.modifiedAsDouble, {});
        shouldThrow(coreProperties.setModifiedAsDouble, coreProperties, 'a');
        shouldThrow(coreProperties.setModifiedAsDouble, {}, 1.0);

        assert.strictEqual(coreProperties.setModifiedAsDouble(1.0), coreProperties);
        assert.ok(coreProperties.modifiedAsDouble() - 1.0 < testUtils.epsilon);
    });

    it('tags and setTags change the title', () => {
        shouldThrow(coreProperties.tags, {});
        shouldThrow(coreProperties.setTags, coreProperties, 1);
        shouldThrow(coreProperties.setTags, {}, 'Foo');

        assert.strictEqual(coreProperties.setTags('Foo'), coreProperties);
        assert.strictEqual(coreProperties.tags(), 'Foo');
    });

    it('categories and setCategories change the title', () => {
        shouldThrow(coreProperties.categories, {});
        shouldThrow(coreProperties.setCategories, coreProperties, 1);
        shouldThrow(coreProperties.setCategories, {}, 'Foo');

        assert.strictEqual(coreProperties.setCategories('Foo'), coreProperties);
        assert.strictEqual(coreProperties.categories(), 'Foo');
    });

    it('comments and setComments change the title', () => {
        shouldThrow(coreProperties.comments, {});
        shouldThrow(coreProperties.setComments, coreProperties, 1);
        shouldThrow(coreProperties.setComments, {}, 'Foo');

        assert.strictEqual(coreProperties.setComments('Foo'), coreProperties);
        assert.strictEqual(coreProperties.comments(), 'Foo');
    });

    it('removeAll unsets all properties', () => {
        shouldThrow(coreProperties.removeAll, {});

        coreProperties.setTitle('Foo');

        assert.strictEqual(coreProperties.removeAll(), coreProperties);
        assert.throws(() => coreProperties.title());
    });
});
